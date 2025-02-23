#include <windows.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <string>
#include <CommCtrl.h>
#include <map>
#include <set>

#pragma comment(lib, "ComCtl32.lib")
#pragma comment(lib, "Ole32.lib")

#define WINDOW_WIDTH 280
#define WINDOW_HEIGHT 130
#define TIMER_ID 1
#define TIMER_INTERVAL 1500
#define ID_CLOSE_BUTTON 1
#define ID_TOGGLE_TOP_BUTTON 2
#define ID_TOGGLE_MIC_BUTTON 3
#define ID_VOLUME_SLIDER 4
#define ID_VOLUME_TEXT 5
#define ID_HOTKEY_BUTTON 6
#define ID_HOTKEY_WINDOW 100
#define ID_HOTKEY_CLOSE_BUTTON 101
#define ID_HOTKEY_SET_BUTTON 102
#define ID_HOTKEY_KEY1 103
#define ID_HOTKEY_KEY2 104
#define ID_HOTKEY_KEY3 105
#define ID_HOTKEY_KEY4 106
#define ID_HOTKEY_INDICATOR 107

HWND hHotkeyWindow = NULL;
HWND hMainWindow = NULL;
std::map<int, UINT> keyButtons = {
    {ID_HOTKEY_KEY1, 0},
    {ID_HOTKEY_KEY2, 0},
    {ID_HOTKEY_KEY3, 0},
    {ID_HOTKEY_KEY4, 0}
};

bool IsMicrophoneMuted() {
    HRESULT hr;
    BOOL isMuted = FALSE;
    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDevice* pDevice = NULL;
    IAudioEndpointVolume* pEndpointVolume = NULL;
    hr = CoInitialize(NULL);
    if (FAILED(hr)) return false;
    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
    if (SUCCEEDED(hr)) {
        hr = pEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &pDevice);
        if (SUCCEEDED(hr)) {
            hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&pEndpointVolume);
            if (SUCCEEDED(hr)) {
                pEndpointVolume->GetMute(&isMuted);
                pEndpointVolume->Release();
            }
            pDevice->Release();
        }
        pEnumerator->Release();
    }
    CoUninitialize();
    return isMuted != FALSE;
}

float GetMicrophoneVolumeLevel() {
    HRESULT hr;
    float volumeLevel = 0.0f;
    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDevice* pDevice = NULL;
    IAudioEndpointVolume* pEndpointVolume = NULL;
    hr = CoInitialize(NULL);
    if (FAILED(hr)) return volumeLevel;
    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
    if (SUCCEEDED(hr)) {
        hr = pEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &pDevice);
        if (SUCCEEDED(hr)) {
            hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&pEndpointVolume);
            if (SUCCEEDED(hr)) {
                pEndpointVolume->GetMasterVolumeLevelScalar(&volumeLevel);
                pEndpointVolume->Release();
            }
            pDevice->Release();
        }
        pEnumerator->Release();
    }
    CoUninitialize();
    return volumeLevel;
}

void SetMicrophoneVolumeLevel(float volumeLevel) {
    HRESULT hr;
    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDevice* pDevice = NULL;
    IAudioEndpointVolume* pEndpointVolume = NULL;
    hr = CoInitialize(NULL);
    if (FAILED(hr)) return;
    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
    if (SUCCEEDED(hr)) {
        hr = pEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &pDevice);
        if (SUCCEEDED(hr)) {
            hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&pEndpointVolume);
            if (SUCCEEDED(hr)) {
                pEndpointVolume->SetMasterVolumeLevelScalar(volumeLevel, NULL);
                pEndpointVolume->Release();
            }
            pDevice->Release();
        }
        pEnumerator->Release();
    }
    CoUninitialize();
}

void ToggleMicrophoneMute() {
    HRESULT hr;
    IMMDeviceEnumerator* pEnumerator = NULL;
    IMMDevice* pDevice = NULL;
    IAudioEndpointVolume* pEndpointVolume = NULL;
    hr = CoInitialize(NULL);
    if (FAILED(hr)) return;
    hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnumerator);
    if (SUCCEEDED(hr)) {
        hr = pEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &pDevice);
        if (SUCCEEDED(hr)) {
            hr = pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, NULL, (void**)&pEndpointVolume);
            if (SUCCEEDED(hr)) {
                BOOL isMuted;
                pEndpointVolume->GetMute(&isMuted);
                pEndpointVolume->SetMute(!isMuted, NULL);
                pEndpointVolume->Release();
            }
            pDevice->Release();
        }
        pEnumerator->Release();
    }
    CoUninitialize();
}

void DrawWindow(HWND hwnd) {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hwnd, &ps);
    RECT rect;
    GetClientRect(hwnd, &rect);
    HBRUSH greenBrush = CreateSolidBrush(RGB(0, 255, 0));
    HBRUSH redBrush = CreateSolidBrush(RGB(255, 0, 0));
    SelectObject(hdc, IsMicrophoneMuted() ? redBrush : greenBrush);
    Rectangle(hdc, 2, 2, rect.right - 2, rect.bottom - 2);
    SetTextColor(hdc, IsMicrophoneMuted() ? RGB(255, 255, 255) : RGB(0, 0, 0));
    SetBkMode(hdc, TRANSPARENT);
    HFONT hFont = CreateFont(24, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, VARIABLE_PITCH, TEXT("Verdana"));
    SelectObject(hdc, hFont);
    std::wstring statusText = IsMicrophoneMuted() ? L"Микрофон выключен" : L"В Эфире";
    rect.top -= 20;
    DrawTextW(hdc, statusText.c_str(), -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    DeleteObject(greenBrush);
    DeleteObject(redBrush);
    DeleteObject(hFont);
    EndPaint(hwnd, &ps);
}

void SetIndicatorColor(HWND hwnd, COLORREF color) {
    HWND hIndicator = GetDlgItem(hwnd, ID_HOTKEY_INDICATOR);
    HDC hdc = GetDC(hIndicator);
    RECT rect;
    GetClientRect(hIndicator, &rect);
    HBRUSH hBrush = CreateSolidBrush(color);
    FillRect(hdc, &rect, hBrush);
    DeleteObject(hBrush);
    ReleaseDC(hIndicator, hdc);
}

LRESULT CALLBACK HotkeyWindowProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static bool waitingForKey = false;
    static int currentButtonID = 0;

    switch (msg) {
    case WM_CREATE: {
        CreateWindowW(L"BUTTON", L"Key 1", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 10, 10, 80, 30, hwnd, (HMENU)ID_HOTKEY_KEY1, NULL, NULL);
        CreateWindowW(L"BUTTON", L"Key 2", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 100, 10, 80, 30, hwnd, (HMENU)ID_HOTKEY_KEY2, NULL, NULL);
        CreateWindowW(L"BUTTON", L"Key 3", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 190, 10, 80, 30, hwnd, (HMENU)ID_HOTKEY_KEY3, NULL, NULL);
        CreateWindowW(L"BUTTON", L"Key 4", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 280, 10, 80, 30, hwnd, (HMENU)ID_HOTKEY_KEY4, NULL, NULL);
        CreateWindowW(L"BUTTON", L"Задать hot-key", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 10, 50, 200, 40, hwnd, (HMENU)ID_HOTKEY_SET_BUTTON, NULL, NULL);
        CreateWindowW(L"BUTTON", L"X", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 220, 50, 40, 40, hwnd, (HMENU)ID_HOTKEY_CLOSE_BUTTON, NULL, NULL);
        CreateWindowW(L"STATIC", NULL, WS_VISIBLE | WS_CHILD | SS_NOTIFY, 270, 50, 55, 40, hwnd, (HMENU)ID_HOTKEY_INDICATOR, NULL, NULL);
        HFONT hFont = CreateFont(16, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, VARIABLE_PITCH, TEXT("Verdana"));
        SendMessage(GetDlgItem(hwnd, ID_HOTKEY_CLOSE_BUTTON), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(hwnd, ID_HOTKEY_SET_BUTTON), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(hwnd, ID_HOTKEY_KEY1), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(hwnd, ID_HOTKEY_KEY2), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(hwnd, ID_HOTKEY_KEY3), WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessage(GetDlgItem(hwnd, ID_HOTKEY_KEY4), WM_SETFONT, (WPARAM)hFont, TRUE);
        break;
    }
    case WM_COMMAND: {
        if (LOWORD(wParam) == ID_HOTKEY_CLOSE_BUTTON) {
            DestroyWindow(hwnd);
            hHotkeyWindow = NULL;
            EnableWindow(hMainWindow, TRUE);
            SetFocus(hMainWindow);
        }
        else if (LOWORD(wParam) >= ID_HOTKEY_KEY1 && LOWORD(wParam) <= ID_HOTKEY_KEY4) {
            waitingForKey = true;
            currentButtonID = LOWORD(wParam);
            SetWindowTextW(GetDlgItem(hwnd, currentButtonID), L"Нажмите клавишу...");
            EnableWindow(GetDlgItem(hwnd, ID_HOTKEY_SET_BUTTON), FALSE);
            for (int i = ID_HOTKEY_KEY1; i <= ID_HOTKEY_KEY4; i++) {
                if (i != currentButtonID) EnableWindow(GetDlgItem(hwnd, i), FALSE);
            }
            SetFocus(hwnd);
        }
        else if (LOWORD(wParam) == ID_HOTKEY_SET_BUTTON) {
            std::set<UINT> uniqueKeys;
            for (const auto& key : keyButtons) {
                if (key.second != 0) uniqueKeys.insert(key.second);
            }
            if (uniqueKeys.size() < 4) {
                SetIndicatorColor(hwnd, RGB(255, 0, 0));
            }
            else {
                UINT modifiers = 0;
                UINT vk = 0;
                if (keyButtons[ID_HOTKEY_KEY1] != 0) modifiers |= MOD_SHIFT;
                if (keyButtons[ID_HOTKEY_KEY2] != 0) modifiers |= MOD_CONTROL;
                if (keyButtons[ID_HOTKEY_KEY3] != 0) modifiers |= MOD_ALT;
                if (keyButtons[ID_HOTKEY_KEY4] != 0) vk = keyButtons[ID_HOTKEY_KEY4];
                if (vk != 0) {
                    if (RegisterHotKey(hMainWindow, 1, modifiers, vk)) {
                        SetIndicatorColor(hwnd, RGB(0, 255, 0));
                    }
                    else {
                        SetIndicatorColor(hwnd, RGB(255, 0, 0));
                    }
                }
            }
        }
        break;
    }
    case WM_KEYDOWN:
    case WM_SYSKEYDOWN: {
        if (waitingForKey && currentButtonID >= ID_HOTKEY_KEY1 && currentButtonID <= ID_HOTKEY_KEY4) {
            UINT vk = (UINT)wParam;
            keyButtons[currentButtonID] = vk;
            wchar_t keyName[256];
            GetKeyNameTextW(lParam, keyName, 256);
            SetWindowTextW(GetDlgItem(hwnd, currentButtonID), keyName);
            waitingForKey = false;
            currentButtonID = 0;
            EnableWindow(GetDlgItem(hwnd, ID_HOTKEY_SET_BUTTON), TRUE);
            for (int i = ID_HOTKEY_KEY1; i <= ID_HOTKEY_KEY4; i++) {
                EnableWindow(GetDlgItem(hwnd, i), TRUE);
            }
        }
        break;
    }
    case WM_NCHITTEST: {
        LRESULT hit = DefWindowProc(hwnd, msg, wParam, lParam);
        if (hit == HTCLIENT) {
            POINT pt = { LOWORD(lParam), HIWORD(lParam) };
            ScreenToClient(hwnd, &pt);
            HWND hChild = ChildWindowFromPoint(hwnd, pt);
            if (hChild == hwnd || hChild == NULL) return HTCAPTION;
        }
        return hit;
    }
    case WM_DESTROY:
        hHotkeyWindow = NULL;
        EnableWindow(hMainWindow, TRUE);
        SetFocus(hMainWindow);
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

void CreateHotkeyWindow(HWND parent) {
    if (hHotkeyWindow) return;
    RECT rect;
    GetWindowRect(GetDlgItem(parent, ID_HOTKEY_BUTTON), &rect);
    WNDCLASSEX wc = { sizeof(WNDCLASSEX), CS_HREDRAW | CS_VREDRAW, HotkeyWindowProc, 0, 0, GetModuleHandle(NULL), NULL, NULL, NULL, NULL, L"HotkeyWindowClass", NULL };
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    RegisterClassEx(&wc);
    hHotkeyWindow = CreateWindowExW(0, wc.lpszClassName, NULL, WS_POPUP | WS_VISIBLE | WS_BORDER, rect.left, rect.bottom, 370, 95, parent, NULL, GetModuleHandle(NULL), NULL);
    EnableWindow(parent, FALSE);
    hMainWindow = parent;
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    static bool isAlwaysOnTop = false;
    static HWND hSlider = NULL;
    static HWND hVolumeText = NULL;
    static bool lastMicrophoneState = IsMicrophoneMuted();

    switch (msg) {
    case WM_CREATE: {
        HFONT hButtonFont = CreateFont(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, VARIABLE_PITCH, TEXT("Verdana"));
        CreateWindowW(L"BUTTON", L"X", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, WINDOW_WIDTH - 30, 10, 22, 22, hwnd, (HMENU)ID_CLOSE_BUTTON, NULL, NULL);
        CreateWindowW(L"BUTTON", L"TOP", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, WINDOW_WIDTH - 102, 10, 74, 22, hwnd, (HMENU)ID_TOGGLE_TOP_BUTTON, NULL, NULL);
        CreateWindowW(L"BUTTON", IsMicrophoneMuted() ? L"Включить" : L"Выключить", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 10, 10, 88, 22, hwnd, (HMENU)ID_TOGGLE_MIC_BUTTON, NULL, NULL);
        CreateWindowW(L"BUTTON", L"Hot key", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 100, 10, 60, 22, hwnd, (HMENU)ID_HOTKEY_BUTTON, NULL, NULL);
        SendMessage(GetDlgItem(hwnd, ID_CLOSE_BUTTON), WM_SETFONT, (WPARAM)hButtonFont, TRUE);
        SendMessage(GetDlgItem(hwnd, ID_TOGGLE_TOP_BUTTON), WM_SETFONT, (WPARAM)hButtonFont, TRUE);
        SendMessage(GetDlgItem(hwnd, ID_TOGGLE_MIC_BUTTON), WM_SETFONT, (WPARAM)hButtonFont, TRUE);
        SendMessage(GetDlgItem(hwnd, ID_HOTKEY_BUTTON), WM_SETFONT, (WPARAM)hButtonFont, TRUE);
        hSlider = CreateWindowW(TRACKBAR_CLASS, NULL, WS_VISIBLE | WS_CHILD | TBS_HORZ | TBS_NOTICKS, 10, WINDOW_HEIGHT - 35, WINDOW_WIDTH - 84, 30, hwnd, (HMENU)ID_VOLUME_SLIDER, NULL, NULL);
        SendMessage(hSlider, TBM_SETRANGE, TRUE, MAKELONG(0, 100));
        SendMessage(hSlider, TBM_SETPOS, TRUE, (LPARAM)(GetMicrophoneVolumeLevel() * 100));
        hVolumeText = CreateWindowW(L"STATIC", L"100%", WS_VISIBLE | WS_CHILD | SS_CENTER, WINDOW_WIDTH - 80, WINDOW_HEIGHT - 35, 70, 30, hwnd, (HMENU)ID_VOLUME_TEXT, NULL, NULL);
        HFONT hFont = CreateFont(16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, VARIABLE_PITCH, TEXT("Verdana"));
        SendMessage(hVolumeText, WM_SETFONT, (WPARAM)hFont, TRUE);
        SetTimer(hwnd, TIMER_ID, TIMER_INTERVAL, NULL);
        break;
    }
    case WM_HOTKEY: {
        if (wParam == 1) ToggleMicrophoneMute();
        SetWindowTextW(GetDlgItem(hwnd, ID_TOGGLE_MIC_BUTTON), IsMicrophoneMuted() ? L"Включить" : L"Выключить");
        InvalidateRect(hwnd, NULL, TRUE);
        break;
    }
    case WM_COMMAND: {
        if (LOWORD(wParam) == ID_CLOSE_BUTTON) PostQuitMessage(0);
        else if (LOWORD(wParam) == ID_TOGGLE_TOP_BUTTON) {
            isAlwaysOnTop = !isAlwaysOnTop;
            SetWindowPos(hwnd, isAlwaysOnTop ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
            SetWindowTextW(GetDlgItem(hwnd, ID_TOGGLE_TOP_BUTTON), isAlwaysOnTop ? L"Normal" : L"TOP");
        }
        else if (LOWORD(wParam) == ID_TOGGLE_MIC_BUTTON) {
            ToggleMicrophoneMute();
            SetWindowTextW(GetDlgItem(hwnd, ID_TOGGLE_MIC_BUTTON), IsMicrophoneMuted() ? L"Включить" : L"Выключить");
            InvalidateRect(hwnd, NULL, TRUE);
        }
        else if (LOWORD(wParam) == ID_HOTKEY_BUTTON) CreateHotkeyWindow(hwnd);
        break;
    }
    case WM_HSCROLL: {
        if ((HWND)lParam == hSlider) {
            int pos = SendMessage(hSlider, TBM_GETPOS, 0, 0);
            SetMicrophoneVolumeLevel(pos / 100.0f);
            wchar_t volumeStr[10];
            swprintf(volumeStr, 10, L"%d%%", pos);
            SetWindowTextW(hVolumeText, volumeStr);
        }
        break;
    }
    case WM_TIMER: {
        float volumeLevel = GetMicrophoneVolumeLevel();
        SendMessage(hSlider, TBM_SETPOS, TRUE, (LPARAM)(volumeLevel * 100));
        int pos = SendMessage(hSlider, TBM_GETPOS, 0, 0);
        wchar_t volumeStr[10];
        swprintf(volumeStr, 10, L"%d%%", pos);
        SetWindowTextW(hVolumeText, volumeStr);
        bool currentMicrophoneState = IsMicrophoneMuted();
        if (currentMicrophoneState != lastMicrophoneState) {
            SetWindowTextW(GetDlgItem(hwnd, ID_TOGGLE_MIC_BUTTON), currentMicrophoneState ? L"Включить" : L"Выключить");
            InvalidateRect(hwnd, NULL, TRUE);
            lastMicrophoneState = currentMicrophoneState;
        }
        break;
    }
    case WM_PAINT:
        DrawWindow(hwnd);
        break;
    case WM_NCHITTEST: {
        LRESULT hit = DefWindowProc(hwnd, msg, wParam, lParam);
        return hit == HTCLIENT ? HTCAPTION : hit;
    }
    case WM_DESTROY:
        UnregisterHotKey(hwnd, 1);
        KillTimer(hwnd, TIMER_ID);
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
    return 0;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow) {
    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_WIN95_CLASSES };
    InitCommonControlsEx(&icc);
    WNDCLASSEX wc = { sizeof(WNDCLASSEX), CS_HREDRAW | CS_VREDRAW, WndProc, 0, 0, hInstance, NULL, NULL, NULL, NULL, L"MicrophoneStatus", NULL };
    RegisterClassEx(&wc);
    HWND hwnd = CreateWindowExW(0, wc.lpszClassName, L"Microphone Status", WS_POPUP | WS_VISIBLE, 100, 100, WINDOW_WIDTH, WINDOW_HEIGHT, NULL, NULL, hInstance, NULL);
    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    UnregisterClass(wc.lpszClassName, wc.hInstance);
    return 0;
}
