// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.IO.Abstractions;
using System.Runtime.CompilerServices;
using Common;
using Microsoft.PowerToys.FilePreviewCommon;
using Microsoft.PowerToys.PreviewHandler.PhotoGeo.Telemetry.Events;
using Microsoft.PowerToys.Telemetry;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace Microsoft.PowerToys.PreviewHandler.PhotoGeo
{
    /// <summary>
    /// Win Form Implementation for PhotoGeo Preview Handler.
    /// </summary>
    public partial class PhotoGeoPreviewHandlerControl : FormHandlerControl
    {
        private static readonly IFileSystem FileSystem = new FileSystem();
        private static readonly IPath Path = FileSystem.Path;
        private static readonly IFile File = FileSystem.File;

        /// <summary>
        /// RichTextBox control to display error messages.
        /// </summary>
        private RichTextBox _infoBar;

        /// <summary>
        /// Extended Browser Control to display HTML content.
        /// </summary>
        private WebView2 _browser;

        /// <summary>
        /// WebView2 Environment
        /// </summary>
        private CoreWebView2Environment _webView2Environment;

        /// <summary>
        /// Name of the virtual host
        /// </summary>
        public const string VirtualHostName = "PowerToysLocalPhotoGeo";

        /// <summary>
        /// URI of the local file saved with the contents
        /// </summary>
        private Uri _localFileURI;

        /// <summary>
        /// True if info bar is displayed, false otherwise.
        /// </summary>
        private bool _infoBarDisplayed;

        /// <summary>
        /// Gets the path of the current assembly.
        /// </summary>
        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = AppContext.BaseDirectory;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        /// <summary>
        /// Represent WebView2 user data folder path.
        /// </summary>
        private string _webView2UserDataFolder = System.Environment.GetEnvironmentVariable("USERPROFILE") +
                                "\\AppData\\LocalLow\\Microsoft\\PowerToys\\PhotoGeoPreview-Temp";

        /// <summary>
        /// Initializes a new instance of the <see cref="PhotoGeoPreviewHandlerControl"/> class.
        /// </summary>
        public PhotoGeoPreviewHandlerControl()
        {
            this.SetBackgroundColor(Settings.BackgroundColor);
        }

        /// <summary>
        /// Disposes the resources used by the control.
        /// </summary>
        /// <param name="disposing">True if called from Dispose, false otherwise.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _browser?.Dispose();
                _browser = null;
                _infoBar?.Dispose();
                _infoBar = null;
            }

            base.Dispose(disposing);
        }

        /// <summary>
        /// Start the preview on the Control.
        /// </summary>
        /// <param name="dataSource">Path to the file.</param>
        public override async void DoPreview<T>(T dataSource)
        {
            // Temporarily disabled for development - uncomment before PR
            // if (global::PowerToys.GPOWrapper.GPOWrapper.GetConfiguredPhotoGeoPreviewEnabledValue() == global::PowerToys.GPOWrapper.GpoRuleConfigured.Disabled)
            // {
            //     // GPO is disabling this utility. Show an error message instead.
            //     _infoBarDisplayed = true;
            //     _infoBar = GetTextBoxControl("This preview handler is disabled by Group Policy.");
            //     Resize += FormResized;
            //     Controls.Add(_infoBar);
            //     base.DoPreview(dataSource);
            //     return;
            // }
            Helper.CleanupTempDir(_webView2UserDataFolder);

            _infoBarDisplayed = false;

            try
            {
                if (!(dataSource is string filePath))
                {
                    throw new ArgumentException($"{nameof(dataSource)} for {nameof(PhotoGeoPreviewHandlerControl)} must be a string but was a '{typeof(T)}'");
                }

                // Extract EXIF GPS data
                (double? latitude, double? longitude) = await ExtractGpsDataAsync(filePath);
                string htmlContent = GenerateHtmlContent(filePath, latitude, longitude);

                _browser = new WebView2()
                {
                    Dock = DockStyle.Fill,
                    DefaultBackgroundColor = Color.Transparent,
                };

                var webView2Options = new CoreWebView2EnvironmentOptions("--block-new-web-contents");
                try
                {
                    _webView2Environment = await CoreWebView2Environment
                        .CreateAsync(userDataFolder: _webView2UserDataFolder, options: webView2Options)
                        .ConfigureAwait(true);

                    Controls.Add(_browser);

                    await _browser.EnsureCoreWebView2Async(_webView2Environment).ConfigureAwait(true);
                    _browser.CoreWebView2.SetVirtualHostNameToFolderMapping(VirtualHostName, AssemblyDirectory, CoreWebView2HostResourceAccessKind.Deny);
                    _browser.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = false;
                    _browser.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
                    _browser.CoreWebView2.Settings.AreDevToolsEnabled = false;
                    _browser.CoreWebView2.Settings.AreHostObjectsAllowed = false;
                    _browser.CoreWebView2.Settings.IsGeneralAutofillEnabled = false;
                    _browser.CoreWebView2.Settings.IsPasswordAutosaveEnabled = false;
                    _browser.CoreWebView2.Settings.IsScriptEnabled = true; // Need JavaScript for map
                    _browser.CoreWebView2.Settings.IsWebMessageEnabled = false;

                    // Allow loading resources for the map
                    _browser.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All);
                    _browser.CoreWebView2.WebResourceRequested += (object sender, CoreWebView2WebResourceRequestedEventArgs e) =>
                    {
                        // Allow local file and external resources (Leaflet CDN, OpenStreetMap tiles)
                        var uri = new Uri(e.Request.Uri);
                        if (uri != _localFileURI &&
                            !uri.Host.Contains("unpkg.com", StringComparison.OrdinalIgnoreCase) &&
                            !uri.Host.Contains("openstreetmap.org", StringComparison.OrdinalIgnoreCase) &&
                            !uri.Scheme.Equals("data", StringComparison.OrdinalIgnoreCase))
                        {
                            e.Response = _browser.CoreWebView2.Environment.CreateWebResourceResponse(null, 403, "Forbidden", null);
                        }
                    };

                    // Save HTML to temp file and navigate
                    string filename = _webView2UserDataFolder + "\\" + Guid.NewGuid().ToString() + ".html";
                    File.WriteAllText(filename, htmlContent);
                    _localFileURI = new Uri(filename);
                    _browser.Source = _localFileURI;
                }
                catch (NullReferenceException)
                {
                }

                try
                {
                    PowerToysTelemetry.Log.WriteEvent(new PhotoGeoFilePreviewed());
                }
                catch
                { // Should not crash if sending telemetry is failing. Ignore the exception.
                }
            }
            catch (Exception ex)
            {
                try
                {
                    PowerToysTelemetry.Log.WriteEvent(new PhotoGeoFilePreviewError { Message = ex.Message });
                }
                catch
                { // Should not crash if sending telemetry is failing. Ignore the exception.
                }

                Controls.Clear();
                _infoBarDisplayed = true;
                _infoBar = GetTextBoxControl("Failed to preview image: " + ex.Message);
                Resize += FormResized;
                Controls.Add(_infoBar);
            }
            finally
            {
                base.DoPreview(dataSource);
            }
        }

        /// <summary>
        /// Generate HTML content for preview.
        /// </summary>
        /// <param name="filePath">Path to the image file.</param>
        /// <param name="latitude">Latitude from EXIF (optional).</param>
        /// <param name="longitude">Longitude from EXIF (optional).</param>
        /// <returns>HTML content string.</returns>
        private string GenerateHtmlContent(string filePath, double? latitude, double? longitude)
        {
            // Convert image to Base64 data URI
            string imageUri = ConvertImageToBase64DataUri(filePath);

            if (latitude.HasValue && longitude.HasValue)
            {
                // HTML with image and map
                return $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <link rel='stylesheet' href='https://unpkg.com/leaflet@1.9.4/dist/leaflet.css' />
    <style>
        body {{
            margin: 0;
            padding: 0;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            display: flex;
            flex-direction: column;
            height: 100vh;
            overflow: hidden;
        }}
        .container {{
            display: flex;
            flex-direction: column;
            height: 100vh;
        }}
        .image-pane {{
            flex: 1;
            display: flex;
            justify-content: center;
            align-items: center;
            background: #f0f0f0;
            overflow: hidden;
        }}
        .image-pane img {{
            max-width: 100%;
            max-height: 100%;
            object-fit: contain;
        }}
        .splitter {{
            height: 5px;
            background: #ccc;
            cursor: ns-resize;
            user-select: none;
        }}
        .splitter:hover {{
            background: #999;
        }}
        .map-pane {{
            flex: 1;
            position: relative;
        }}
        #map {{
            width: 100%;
            height: 100%;
        }}
    </style>
</head>
<body>
    <div class='container'>
        <div class='image-pane' id='imagePane'>
            <img src='{imageUri}' alt='Photo' />
        </div>
        <div class='splitter' id='splitter'></div>
        <div class='map-pane'>
            <div id='map'></div>
        </div>
    </div>
    <script src='https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'></script>
    <script>
        // Initialize map
        var map = L.map('map').setView([{latitude.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}, {longitude.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}], 13);

        L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
            attribution: '&copy; <a href=""https://www.openstreetmap.org/copyright"">OpenStreetMap</a> contributors'
        }}).addTo(map);

        L.marker([{latitude.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}, {longitude.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}]).addTo(map)
            .bindPopup('Photo location')
            .openPopup();

        // Splitter drag functionality
        const splitter = document.getElementById('splitter');
        const imagePane = document.getElementById('imagePane');
        let isDragging = false;

        splitter.addEventListener('mousedown', () => {{
            isDragging = true;
        }});

        document.addEventListener('mouseup', () => {{
            isDragging = false;
        }});

        document.addEventListener('mousemove', (e) => {{
            if (isDragging) {{
                const container = document.querySelector('.container');
                const containerRect = container.getBoundingClientRect();
                const percentage = ((e.clientY - containerRect.top) / containerRect.height) * 100;

                if (percentage > 10 && percentage < 90) {{
                    imagePane.style.flex = percentage;
                    document.querySelector('.map-pane').style.flex = 100 - percentage;
                }}
            }}
        }});
    </script>
</body>
</html>";
            }
            else
            {
                // HTML with image only (no GPS data)
                return $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <style>
        body {{
            margin: 0;
            padding: 0;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            background: #f0f0f0;
        }}
        img {{
            max-width: 100%;
            max-height: 100%;
            object-fit: contain;
        }}
        .no-gps {{
            position: absolute;
            top: 10px;
            left: 10px;
            background: #fff;
            padding: 10px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.2);
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            font-size: 14px;
        }}
    </style>
</head>
<body>
    <div class='no-gps'>No GPS data found in this image</div>
    <img src='{imageUri}' alt='Photo' />
</body>
</html>";
            }
        }

        /// <summary>
        /// Gets a textbox control.
        /// </summary>
        /// <param name="message">Message to be displayed in textbox.</param>
        /// <returns>An object of type <see cref="RichTextBox"/>.</returns>
        private RichTextBox GetTextBoxControl(string message)
        {
            RichTextBox richTextBox = new RichTextBox
            {
                Text = message,
                BackColor = Color.LightYellow,
                Multiline = true,
                Dock = DockStyle.Top,
                ReadOnly = true,
            };
            richTextBox.ContentsResized += RTBContentsResized;
            richTextBox.ScrollBars = RichTextBoxScrollBars.None;
            richTextBox.BorderStyle = BorderStyle.None;

            return richTextBox;
        }

        /// <summary>
        /// Callback when RichTextBox is resized.
        /// </summary>
        private void RTBContentsResized(object sender, ContentsResizedEventArgs e)
        {
            RichTextBox richTextBox = (RichTextBox)sender;
            richTextBox.Height = e.NewRectangle.Height + 5;
        }

        /// <summary>
        /// Callback when form is resized.
        /// </summary>
        private void FormResized(object sender, EventArgs e)
        {
            if (_infoBarDisplayed)
            {
                _infoBar.Width = Width;
            }
        }

        /// <summary>
        /// Extract GPS data from image EXIF metadata.
        /// </summary>
        /// <param name="filePath">Path to the image file.</param>
        /// <returns>Tuple of latitude and longitude, or null if no GPS data found.</returns>
        private async Task<(double? Latitude, double? Longitude)> ExtractGpsDataAsync(string filePath)
        {
            try
            {
                var file = await Windows.Storage.StorageFile.GetFileFromPathAsync(filePath);
                var properties = await file.Properties.GetImagePropertiesAsync();

                if (properties.Latitude.HasValue && properties.Longitude.HasValue)
                {
                    return (properties.Latitude.Value, properties.Longitude.Value);
                }
            }
            catch (Exception)
            {
                // Failed to extract GPS data, return null
            }

            return (null, null);
        }

        /// <summary>
        /// Convert image file to Base64 data URI.
        /// </summary>
        /// <param name="filePath">Path to the image file.</param>
        /// <returns>Base64 encoded data URI string.</returns>
        private string ConvertImageToBase64DataUri(string filePath)
        {
            try
            {
                byte[] imageBytes = File.ReadAllBytes(filePath);
                string base64String = Convert.ToBase64String(imageBytes);

                // Determine MIME type from file extension
                string extension = Path.GetExtension(filePath).ToUpperInvariant();
                string mimeType = extension switch
                {
                    ".JPG" or ".JPEG" => "image/jpeg",
                    ".PNG" => "image/png",
                    ".GIF" => "image/gif",
                    ".BMP" => "image/bmp",
                    ".WEBP" => "image/webp",
                    _ => "image/jpeg",
                };

                return $"data:{mimeType};base64,{base64String}";
            }
            catch (Exception)
            {
                // Return empty data URI on error
                return "data:image/png;base64,";
            }
        }
    }
}
