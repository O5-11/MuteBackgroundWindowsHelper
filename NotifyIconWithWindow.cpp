
#pragma region 头文件及宏定义
#include "framework.h"
#include "NotifyIconWithWindow.h"
#include "shellapi.h"
#include "c:\Head\Head.h"
#include "vector"
#include <comdef.h>
#include <mmdeviceapi.h> 
#include <endpointvolume.h>
#include <audioclient.h>
#include <audiopolicy.h>
#include "cJSON.h"

#include <atlbase.h>

#include "Resource.h"

#pragma comment(lib,"winmm.lib")

using namespace std;


#define WM_SYSTEMTRAY      WM_USER+100  //自定义消息，用于弹出系统托盘

#define WM_START    WM_USER+101     //启动
#define WM_STOP     WM_USER+102     //停止
#define WM_QUIT     WM_USER+103     //退出
#define WM_NONE     WM_USER+104     //无效
#define WM_SHOW     WM_USER+105     //弹出

#define SETTINGFILENAME "Background Mute Windows List.json"

#define MAX_LOADSTRING 100

#pragma endregion

#pragma region 自定义数据结构及全局变量

typedef struct Data
{
    char* Name;
    ULONG pID;
    bool FindState;
    ISimpleAudioVolume* Controler;
}Data;

vector<Data>List;

bool SINGLE = false;

#pragma endregion

#pragma region 自定义功能函数

ULONG GetProcessIdFromName(const char* name)
{
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnapshot == INVALID_HANDLE_VALUE)
    {
        MBX("创建进程快照失败");
        return -1;
    }

    PROCESSENTRY32 processer;
    processer.dwSize = sizeof(PROCESSENTRY32);

    int flag = Process32First(hSnapshot, &processer);
    while (flag != 0)
    {
        _bstr_t processName(processer.szExeFile);  //WCHAR字符串转换成CHAR字符串
        if (strcmp(processName, name) == 0)
        {
            return processer.th32ProcessID;        //返回进程ID
        }
        flag = Process32Next(hSnapshot, &processer);
    }

    CloseHandle(hSnapshot);
    return -2;
};

class SingleVolume
{
public:
    ISimpleAudioVolume* pSimplevol = NULL;

    SingleVolume()
    {
        CoInitialize(0);
    };

    ~SingleVolume()
    {
        CoUninitialize();
    };

    ISimpleAudioVolume* GetTargetProcessVolumeControl(ULONG TargetpId)
    {
        HRESULT hr = S_OK;
        IMMDeviceCollection* pMultiDevice = NULL;
        IMMDevice* pDevice = NULL;
        IAudioSessionEnumerator* pSessionEnum = NULL;
        IAudioSessionManager2* pASManager = NULL;
        IMMDeviceEnumerator* m_pEnumerator = NULL;
        const IID IID_ISimpleAudioVolume = __uuidof(ISimpleAudioVolume);
        const IID IID_IAudioSessionControl2 = __uuidof(IAudioSessionControl2);
        GUID m_guidMyContext;

        hr = CoCreateGuid(&m_guidMyContext);
        if (FAILED(hr))
            return FALSE;
        // Get enumerator for audio endpoint devices.  
        hr = CoCreateInstance(__uuidof(MMDeviceEnumerator),
            NULL, CLSCTX_ALL,
            __uuidof(IMMDeviceEnumerator),
            (void**)&m_pEnumerator);
        if (FAILED(hr))
            return FALSE;

        /*if (IsMixer)
        {
            hr = m_pEnumerator->EnumAudioEndpoints(eRender,DEVICE_STATE_ACTIVE, &pMultiDevice);
        }
        else
        {
            hr = m_pEnumerator->EnumAudioEndpoints(eCapture,DEVICE_STATE_ACTIVE, &pMultiDevice);
        } */
        hr = m_pEnumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pMultiDevice);
        if (FAILED(hr))
            return FALSE;

        UINT deviceCount = 0;
        hr = pMultiDevice->GetCount(&deviceCount);
        if (FAILED(hr))
            return FALSE;

        for (UINT ii = 0; ii < deviceCount; ii++)
        {
            pDevice = NULL;
            hr = pMultiDevice->Item(ii, &pDevice);
            if (FAILED(hr))
                return FALSE;
            hr = pDevice->Activate(__uuidof(IAudioSessionManager), CLSCTX_ALL, NULL, (void**)&pASManager);

            if (FAILED(hr))
                return FALSE;
            hr = pASManager->GetSessionEnumerator(&pSessionEnum);
            if (FAILED(hr))
                return FALSE;
            int nCount;
            hr = pSessionEnum->GetCount(&nCount);

            for (int i = 0; i < nCount; i++)
            {
                IAudioSessionControl* pSessionCtrl;
                hr = pSessionEnum->GetSession(i, &pSessionCtrl);
                if (FAILED(hr))
                    continue;
                IAudioSessionControl2* pSessionCtrl2;
                hr = pSessionCtrl->QueryInterface(IID_IAudioSessionControl2, (void**)&pSessionCtrl2);
                if (FAILED(hr))
                    continue;
                ULONG pid;
                hr = pSessionCtrl2->GetProcessId(&pid);
                if (FAILED(hr))
                    continue;

                hr = pSessionCtrl2->QueryInterface(IID_ISimpleAudioVolume, (void**)&pSimplevol);
                if (FAILED(hr))
                    continue;

                if (pid == TargetpId)
                {
                    m_pEnumerator->Release();
                    return pSimplevol;
                    /*pSimplevol->SetMasterVolume((float)dwVolume/100, NULL);

                    if (dwVolume == 0)
                        pSimplevol->SetMute(true, NULL);*/
                }
            }
        }
        m_pEnumerator->Release();
        return pSimplevol;
    }

    bool IsMuted(ISimpleAudioVolume* ISAV)
    {
        BOOL state = false;
        if (ISAV != NULL)
            ISAV->GetMute(&state);
        return state;
    };

    bool SetMute(ISimpleAudioVolume* ISAV)
    {
        if (ISAV != NULL)
        {
            ISAV->SetMasterVolume(0.0f, NULL);
            ISAV->SetMute(true, NULL);
            return true;
        }
        return false;
    }

    bool UnMute(ISimpleAudioVolume* ISAV)
    {
        if (ISAV != NULL)
        {
            ISAV->SetMasterVolume(1.0f, NULL);
            ISAV->SetMute(false, NULL);
            return true;
        }
        return false;
    }
};

//取消掉cjson返回的字符串两端冒号
void LeftMH(char* str)
{
    int size = strlen(str);
    str[0] = 0;
    str[size - 1] = 0;
    for (int i = 1; i < size; i++)
    {
        str[i - 1] = str[i];
    }
}

//获取文件长度
size_t get_file_size(const char* filepath)
{

    if (NULL == filepath)
        return 0;
    struct stat filestat;
    memset(&filestat, 0, sizeof(struct stat));
    /*获取文件信息*/
    if (0 == stat(filepath, &filestat))
        return filestat.st_size;
    else
        return 0;
}

void ReadJson(const char* FileName)
{
    if (_access(FileName, 0) != 0)
    {
        MBX("配置 \"Background Mute Windows List.json\" 文件不存在"); return;
    }
    size_t filesize = get_file_size(FileName);

    FILE* f = fopen(FileName, "rb");
    if (f == NULL)
    {
        MBX("配置文件正被占用"); return;
    }
    char* buff = NewStr(filesize + 1);
    fread(buff, filesize, 1, f);
    fclose(f);
    
    cJSON* root = NULL;
    root = cJSON_Parse(buff);
    if (!root) { MBX("json数据格式不合法，请检查"); delete[]buff; return; }

    //清空队列
    vector <Data>().swap(List);

    cJSON* head = NULL;
    head = cJSON_GetObjectItem(root,"ProcessName");
    if (!root) { MBX("ProcessName 为空，请检查"); delete[]buff; return; }

    cJSON* item = NULL;
    int size = cJSON_GetArraySize(head);

    for (int i = 0; i < size; i++)
    {
        Data data;
        item = cJSON_GetArrayItem(head, i);
        char* nameptr = cJSON_Print(item);
        data.Name = NewStr(strlen(nameptr));
        strcpy(data.Name, nameptr);
        LeftMH(data.Name);
        data.FindState = false;
        data.pID = 0;
        data.Controler = NULL;
        List.push_back(data);
    }

    delete[]buff;
    cJSON_Delete(root);

    return;
};

void EnumpId(LPVOID args)
{
    SingleVolume sv;
    while (SINGLE)
    {
        if (List.size() >= 1)
        {
            for (int i = 0; SINGLE && List.size() >= 1;)
            {
                Sleep(1);

                if (List[i].FindState == false)
                {
                    List[i].pID = GetProcessIdFromName(List[i].Name);
                    if (List[i].pID > 0)
                    {
                        List[i].FindState = true;
                        List[i].Controler = sv.GetTargetProcessVolumeControl(List[i].pID);
                    }
                }

                i++;
                if (i >= List.size())i = 0;
            }
        }

        Sleep(1);
    }

    return;
};

void MainFunc(LPVOID args)
{
    SingleVolume sv;
    while (SINGLE)
    {
        if (List.size() >= 1)
        {
            for (int i = 0; SINGLE && List.size() >= 1;)
            {
                Sleep(1);
                ULONG forgepId = 0;
                GetWindowThreadProcessId(GetForegroundWindow(), &forgepId);

                if (forgepId != 0 && forgepId == List[i].pID)
                {
                    if (sv.IsMuted(List[i].Controler))
                        sv.UnMute(List[i].Controler);
                }
                else
                {
                    if (!sv.IsMuted(List[i].Controler))
                        sv.SetMute(List[i].Controler);
                }

                i++;
                if (i >= List.size())i = 0;
            }
        }

        Sleep(1);
    }
};

#pragma endregion

#pragma region VS自动生成模板_变量及主函数
// 全局变量:
HINSTANCE hInst;                                // 当前实例
WCHAR szTitle[MAX_LOADSTRING];                  // 标题栏文本
WCHAR szWindowClass[MAX_LOADSTRING];            // 主窗口类名

// 此代码模块中包含的函数的前向声明:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR    lpCmdLine,
    _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: 在此处放置代码。

    // 初始化全局字符串
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_NOTIFYICONWITHWINDOW, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    nCmdShow = SW_HIDE;
    // 执行应用程序初始化:
    if (!InitInstance(hInstance, nCmdShow))
    {
        return FALSE;
    }

    HWND mainWindow = FindWindowW(szWindowClass, szTitle);
    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_NOTIFYICONWITHWINDOW));

    MSG msg;

    NOTIFYICONDATA NotifyIcon;//系统托盘类
    NotifyIcon.cbSize = sizeof(NOTIFYICONDATA);
    //NotifyIcon.hIcon = (HICON)LoadImageA(hInstance, IDI_SMALL, IMAGE_ICON, 0, 0, LR_LOADFROMFILE);
    NotifyIcon.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SMALL));
    NotifyIcon.hWnd = mainWindow;
    NotifyIcon.uID = NULL;
    lstrcpy(NotifyIcon.szTip, _T("窗口助手喵"));
    NotifyIcon.uCallbackMessage = WM_SYSTEMTRAY;
    NotifyIcon.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    Shell_NotifyIcon(NIM_ADD, &NotifyIcon);   //添加系统托盘

    //ULONG pID = GetProcessIdFromName("YM.exe");
    //char* str = new char[16]; CLC(str,16);
    //itoa(pID, str, 10);
    //SingleVolume sv;
    //sv.pSimplevol = sv.GetTargetProcessVolumeControl(pID);
    //sv.SetMute(sv.pSimplevol);
    //MessageBoxA(mainWindow, str, "TEST", MB_OK);
    //delete[]str;

    // 主消息循环:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int)msg.wParam;
};
#pragma endregion

#pragma region VS自动生成模板_暂时无用
// “关于”框的消息处理程序。
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

//
//  函数: MyRegisterClass()
//
//  目标: 注册窗口类。
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_NOTIFYICONWITHWINDOW));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_NOTIFYICONWITHWINDOW);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   函数: InitInstance(HINSTANCE, int)
//
//   目标: 保存实例句柄并创建主窗口
//
//   注释:
//
//        在此函数中，我们在全局变量中保存实例句柄并
//        创建和显示主程序窗口。
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // 将实例句柄存储在全局变量中

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  函数: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  目标: 处理主窗口的消息。
//
//  WM_COMMAND  - 处理应用程序菜单
//  WM_PAINT    - 绘制主窗口
//  WM_DESTROY  - 发送退出消息并返回
//
//
#pragma endregion

#pragma region 消息处理主函数
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    HMENU hMenu = NULL;
    SimpleThread ST;
    HANDLE hThreads[2] = { NULL,NULL };

    switch (message)
    {
    case WM_SYSTEMTRAY:
    {
        switch (lParam)
        {
        case WM_LBUTTONDOWN:
        {
            ShowWindow(hWnd, SW_SHOW);
        }break;

        case WM_RBUTTONDOWN:
        {
            POINT pt;
            GetCursorPos(&pt);
            hMenu = CreatePopupMenu();

            InsertMenuA(hMenu, -1, MF_STRING, WM_NONE, "窗口助手喵");
            InsertMenuA(hMenu, -1, MF_SEPARATOR, WM_NONE, "-");
            InsertMenuA(hMenu, -1, MF_BYPOSITION, WM_START, "启动");
            InsertMenuA(hMenu, -1, MF_BYPOSITION, WM_STOP, "停止");
            AppendMenuA(hMenu, MF_STRING, WM_QUIT, "退出");

            SetForegroundWindow(hWnd);
            TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, NULL, hWnd, NULL);

            if (PostMessage(WM_NULL, 0, 0, 0))
            {//解决托盘菜单不消失
                break;
            }
            else
                DestroyMenu(hMenu);
        }break;

        default:
        {

        }
        break;
        }
    }break;
    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        // 分析菜单选择:
        switch (wmId)
        {
        case IDM_ABOUT:
            DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
            break;
        case IDM_EXIT:
            DestroyWindow(hWnd);
            break;

        case WM_QUIT:
        {
            if(hThreads[0] != NULL && hThreads[1] != NULL)
                WaitForMultipleObjects(2, hThreads, true, INFINITE);
            
            PostQuitMessage(0);
        }break;

        case WM_START:
        {
            if (SINGLE == false)
            {
                SINGLE = true;
                ReadJson(SETTINGFILENAME);
                hThreads[0] = ST.StartThread(EnumpId, NULL);
                hThreads[1] = ST.StartThread(MainFunc, NULL);
            }
        }break;

        case WM_STOP:
        {
            SINGLE = false;
            if (hThreads[0] != NULL && hThreads[1] != NULL)
            {
                WaitForMultipleObjects(2, hThreads, true, INFINITE);
                hThreads[0] = NULL;
                hThreads[1] = NULL;
            }
        }break;

        case WM_SHOW:
        {
            ShowWindow(hWnd,SW_SHOW);
            SetForegroundWindow(hWnd);
        }break;

        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
    }
    break;
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        // TODO: 在此处添加使用 hdc 的任何绘图代码...
        EndPaint(hWnd, &ps);
    }
    break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

#pragma endregion
