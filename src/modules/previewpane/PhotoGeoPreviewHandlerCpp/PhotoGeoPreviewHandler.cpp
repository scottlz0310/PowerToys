#include "pch.h"
#include "PhotoGeoPreviewHandler.h"
#include "../powerpreview/powerpreviewConstants.h"

#include <shellapi.h>
#include <Shlwapi.h>
#include <string>

#include <common/interop/shared_constants.h>
#include <common/logger/logger.h>
#include <common/SettingsAPI/settings_helpers.h>
#include <common/utils/process_path.h>
#include <common/Themes/windows_colors.h>

extern HINSTANCE g_hInst;
extern long g_cDllRef;

PhotoGeoPreviewHandler::PhotoGeoPreviewHandler() :
    m_cRef(1), m_hwndParent(NULL), m_rcParent(), m_punkSite(NULL), m_process(NULL)
{
    m_resizeEvent = CreateEvent(nullptr, false, false, CommonSharedConstants::PHOTOGEO_PREVIEW_RESIZE_EVENT);

    std::filesystem::path logFilePath(PTSettingsHelper::get_local_low_folder_location());
    logFilePath.append(LogSettings::photoGeoPrevLogPath);
    Logger::init(LogSettings::photoGeoPrevLoggerName, logFilePath.wstring(), PTSettingsHelper::get_log_settings_file_location());

    InterlockedIncrement(&g_cDllRef);
}

PhotoGeoPreviewHandler::~PhotoGeoPreviewHandler()
{
    InterlockedDecrement(&g_cDllRef);
}

#pragma region IUnknown

IFACEMETHODIMP PhotoGeoPreviewHandler::QueryInterface(REFIID riid, void** ppv)
{
    static const QITAB qit[] = {
        QITABENT(PhotoGeoPreviewHandler, IPreviewHandler),
        QITABENT(PhotoGeoPreviewHandler, IInitializeWithFile),
        QITABENT(PhotoGeoPreviewHandler, IPreviewHandlerVisuals),
        QITABENT(PhotoGeoPreviewHandler, IOleWindow),
        QITABENT(PhotoGeoPreviewHandler, IObjectWithSite),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG)
PhotoGeoPreviewHandler::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG)
PhotoGeoPreviewHandler::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }
    return cRef;
}

#pragma endregion

#pragma region IInitializationWithFile

IFACEMETHODIMP PhotoGeoPreviewHandler::Initialize(LPCWSTR pszFilePath, DWORD grfMode)
{
    m_filePath = pszFilePath;
    return S_OK;
}

#pragma endregion

#pragma region IPreviewHandler

IFACEMETHODIMP PhotoGeoPreviewHandler::SetWindow(HWND hwnd, const RECT* prc)
{
    if (hwnd && prc)
    {
        m_hwndParent = hwnd;
        m_rcParent = *prc;
    }
    return S_OK;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::SetFocus()
{
    return S_OK;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::QueryFocus(HWND* phwnd)
{
    HRESULT hr = E_INVALIDARG;
    if (phwnd)
    {
        *phwnd = ::GetFocus();
        if (*phwnd)
        {
            hr = S_OK;
        }
        else
        {
            hr = HRESULT_FROM_WIN32(GetLastError());
        }
    }
    return hr;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::TranslateAccelerator(MSG* pmsg)
{
    HRESULT hr = S_FALSE;
    IPreviewHandlerFrame* pFrame = NULL;
    if (m_punkSite && SUCCEEDED(m_punkSite->QueryInterface(&pFrame)))
    {
        hr = pFrame->TranslateAccelerator(pmsg);

        pFrame->Release();
    }
    return hr;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::SetRect(const RECT* prc)
{
    HRESULT hr = E_INVALIDARG;
    if (prc != NULL)
    {
        if (m_rcParent.left == 0 && m_rcParent.top == 0 && m_rcParent.right == 0 && m_rcParent.bottom == 0 && (prc->left != 0 || prc->top != 0 || prc->right != 0 || prc->bottom != 0))
        {
            // PhotoGeoPreviewHandler position first initialisation, do the first preview
            m_rcParent = *prc;
            DoPreview();
        }
        if (!m_resizeEvent)
        {
            Logger::error(L"Failed to create resize event for PhotoGeoPreviewHandler");
        }
        else
        {
            if (m_rcParent.right != prc->right || m_rcParent.left != prc->left || m_rcParent.top != prc->top || m_rcParent.bottom != prc->bottom)
            {
                if (!SetEvent(m_resizeEvent))
                {
                    Logger::error(L"Failed to signal resize event for PhotoGeoPreviewHandler");
                }
            }
        }
        m_rcParent = *prc;
        hr = S_OK;
    }
    return hr;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::DoPreview()
{
    try
    {
        if (m_hwndParent == NULL || (m_rcParent.left == 0 && m_rcParent.top == 0 && m_rcParent.right == 0 && m_rcParent.bottom == 0))
        {
            // Postponing Start PhotoGeoPreviewHandler.exe, parent and position not yet initialized. Preview will be done after initialisation.
            return S_OK;
        }
        Logger::info(L"Starting PhotoGeoPreviewHandler.exe");

        STARTUPINFO info = { sizeof(info) };
        std::wstring cmdLine{ L"\"" + m_filePath + L"\"" };
        cmdLine += L" ";
        std::wostringstream ss;
        ss << std::hex << m_hwndParent;

        cmdLine += ss.str();
        cmdLine += L" ";
        cmdLine += std::to_wstring(m_rcParent.left);
        cmdLine += L" ";
        cmdLine += std::to_wstring(m_rcParent.right);
        cmdLine += L" ";
        cmdLine += std::to_wstring(m_rcParent.top);
        cmdLine += L" ";
        cmdLine += std::to_wstring(m_rcParent.bottom);
        std::wstring appPath = get_module_folderpath(g_hInst) + L"\\PowerToys.PhotoGeoPreviewHandler.exe";

        SHELLEXECUTEINFO sei{ sizeof(sei) };
        sei.fMask = { SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI };
        sei.lpFile = appPath.c_str();
        sei.lpParameters = cmdLine.c_str();
        sei.nShow = SW_SHOWDEFAULT;
        ShellExecuteEx(&sei);

        // Prevent to leak processes: preview is called multiple times when minimizing and restoring Explorer window
        if (m_process)
        {
            TerminateProcess(m_process, 0);
        }

        m_process = sei.hProcess;
    }
    catch (std::exception& e)
    {
        std::wstring errorMessage = std::wstring{ winrt::to_hstring(e.what()) };
        Logger::error(L"Failed to start PhotoGeoPreviewHandler.exe. Error: {}", errorMessage);
    }

    return S_OK;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::Unload()
{
    Logger::info(L"Unload and terminate .exe");

    m_hwndParent = NULL;
    TerminateProcess(m_process, 0);
    return S_OK;
}

#pragma endregion

#pragma region IPreviewHandlerVisuals

IFACEMETHODIMP PhotoGeoPreviewHandler::SetBackgroundColor(COLORREF color)
{
    HBRUSH brush = CreateSolidBrush(WindowsColors::is_dark_mode() ? powerpreviewConstants::DARK_THEME_COLOR : powerpreviewConstants::LIGHT_THEME_COLOR);
    SetClassLongPtr(m_hwndParent, GCLP_HBRBACKGROUND, reinterpret_cast<LONG_PTR>(brush));
    return S_OK;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::SetFont(const LOGFONTW* plf)
{
    return S_OK;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::SetTextColor(COLORREF color)
{
    return S_OK;
}

#pragma endregion

#pragma region IOleWindow

IFACEMETHODIMP PhotoGeoPreviewHandler::GetWindow(HWND* phwnd)
{
    HRESULT hr = E_INVALIDARG;
    if (phwnd)
    {
        *phwnd = m_hwndParent;
        hr = S_OK;
    }
    return hr;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::ContextSensitiveHelp(BOOL fEnterMode)
{
    return E_NOTIMPL;
}

#pragma endregion

#pragma region IObjectWithSite

IFACEMETHODIMP PhotoGeoPreviewHandler::SetSite(IUnknown* punkSite)
{
    if (m_punkSite)
    {
        m_punkSite->Release();
        m_punkSite = NULL;
    }
    return punkSite ? punkSite->QueryInterface(&m_punkSite) : S_OK;
}

IFACEMETHODIMP PhotoGeoPreviewHandler::GetSite(REFIID riid, void** ppv)
{
    *ppv = NULL;
    return m_punkSite ? m_punkSite->QueryInterface(riid, ppv) : E_FAIL;
}

#pragma endregion

#pragma region Helper Functions

#pragma endregion
