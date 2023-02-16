// ScheduledTask.cpp : This file contains the 'main' function. Program execution begins and ends there.
//
#define _WIN32_DCOM

#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <wincred.h>
//  Include the task header file.
#include <taskschd.h>
#include <string>
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")

int main(int argc, wchar_t* argv[])
{
    if (2 > argc)
    {
        printf("Username to execute as was not specified");

        return 1;
    }
    
    wchar_t* username = argv[1];
    printf("Creating a task as %s", username);


    //  Initialize COM.
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr))
    {
        printf("\nCoInitializeEx failed: %x", hr);
        return 1;
    }

    //  Set general COM security levels.
    hr = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL);

    if (FAILED(hr))
    {
        printf("\nCoInitializeSecurity failed: %x", hr);
        CoUninitialize();
        return 1;
    }

    //  ------------------------------------------------------
    //  Create a name for the task.
    LPCWSTR wszTaskName = L"AdSim Scheduled Task";

    //  Get the windows directory and set the path to notepad.exe.
    size_t requiredSize;
    wchar_t* wstrExecutablePath;
    _wgetenv_s(&requiredSize, NULL, 0, L"WINDIR");
    wstrExecutablePath = (wchar_t*)malloc(requiredSize * sizeof(wchar_t));
    _wgetenv_s(&requiredSize, wstrExecutablePath, requiredSize, L"WINDIR");

    //wstrExecutablePath += L"\\SYSTEM32\\NOTEPAD.EXE";

    std::wstring wstringExecutablePath(wstrExecutablePath);
    wstringExecutablePath += std::wstring(L"\\SYSTEM32\\NOTEPAD.EXE");

    //  Create an instance of the Task Service. 
    ITaskService* pService = NULL;
    hr = CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService);
    if (FAILED(hr))
    {
        printf("Failed to create an instance of ITaskService: %x", hr);
        CoUninitialize();
        return 1;
    }

    //  Connect to the task service.
    hr = pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
    if (FAILED(hr))
    {
        printf("ITaskService::Connect failed: %x", hr);
        pService->Release();
        CoUninitialize();
        return 1;
    }

    //  Get the pointer to the root task folder.  
    //  This folder will hold the new task that is registered.
    ITaskFolder* pRootFolder = NULL;
    hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
    if (FAILED(hr))
    {
        printf("Cannot get Root Folder pointer: %x", hr);
        pService->Release();
        CoUninitialize();
        return 1;
    }

    //  If the same task exists, remove it.
    pRootFolder->DeleteTask(_bstr_t(wszTaskName), 0);

    //  Create the task builder object to create the task.
    ITaskDefinition* pTask = NULL;
    hr = pService->NewTask(0, &pTask);

    pService->Release();  // COM clean up.  Pointer is no longer used.
    if (FAILED(hr))
    {
        printf("Failed to create a task definition: %x", hr);
        pRootFolder->Release();
        CoUninitialize();
        return 1;
    }

    //  ------------------------------------------------------
    //  Get the registration info for setting the identification.
    IRegistrationInfo* pRegInfo = NULL;
    hr = pTask->get_RegistrationInfo(&pRegInfo);
    if (FAILED(hr))
    {
        printf("\nCannot get identification pointer: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pRegInfo->put_Author(SysAllocString(L"NetSPI"));
    pRegInfo->Release();  // COM clean up.  Pointer is no longer used.
    if (FAILED(hr))
    {
        printf("\nCannot put identification info: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ////////////////////////////////////////////////////////////////////////////////
    ///
    ////////////////////////////////////////////////////////////////////////////////

    //  Create the principal for the task
    IPrincipal* pPrincipal = NULL;
    hr = pTask->get_Principal(&pPrincipal);
    if (FAILED(hr))
    {
        printf("\nCannot get principal pointer: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    

    //  Set up principal information: 
    hr = pPrincipal->put_Id(_bstr_t(username));
    if (FAILED(hr))
        printf("\nCannot put the principal ID: %x", hr);

    hr = pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
    if (FAILED(hr))
        printf("\nCannot put principal logon type: %x", hr);

    //  Run the task with the least privileges (LUA) 
    hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);
    pPrincipal->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put principal run level: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ////////////////////////////////////////////////////////////////////////////////
    ///
    ////////////////////////////////////////////////////////////////////////////////

    ITaskSettings* pSettings = NULL;
    hr = pTask->get_Settings(&pSettings);
    if (FAILED(hr))
    {
        printf("\nCannot get settings pointer: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    //  Set setting values for the task.
    hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
    pSettings->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put setting info: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ////////////////////////////////////////////////////////////////////////////////
    ///
    ////////////////////////////////////////////////////////////////////////////////

    ITriggerCollection* pTriggerCollection = NULL;
    hr = pTask->get_Triggers(&pTriggerCollection);
    if (FAILED(hr))
    {
        printf("\nCannot get trigger collection: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ITrigger* pTrigger = NULL;
    hr = pTriggerCollection->Create(TASK_TRIGGER_REGISTRATION, &pTrigger);
    pTriggerCollection->Release();
    if (FAILED(hr))
    {
        printf("\nCannot create a registration trigger: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IRegistrationTrigger* pRegistrationTrigger = NULL;
    hr = pTrigger->QueryInterface(IID_IRegistrationTrigger, (void**)&pRegistrationTrigger);
    pTrigger->Release();
    if (FAILED(hr))
    {
        printf("\nQueryInterface call failed on IRegistrationTrigger: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    hr = pRegistrationTrigger->put_Id(_bstr_t(L"OneTimeTrigger"));
    if (FAILED(hr))
    {
        printf("\nCannot put trigger ID: %x", hr);
    }

    //  Define the delay for the registration trigger.
    hr = pRegistrationTrigger->put_Delay(_bstr_t(L"PT30S"));
    pRegistrationTrigger->Release();
    if (FAILED(hr))
    {
        printf("\nCannot put registration trigger delay: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    ////////////////////////////////////////////////////////////////////////////////
    ///
    ////////////////////////////////////////////////////////////////////////////////

    IActionCollection* pActionCollection = NULL;
    hr = pTask->get_Actions(&pActionCollection);
    if (FAILED(hr))
    {
        printf("\nCannot get trigger collection: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    //  Create the action, specifying that it is an executable action.
    IAction* pAction = NULL;
    hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
    pActionCollection->Release();
    if (FAILED(hr))
    {
        printf("\nCannot create the action: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    IExecAction* pExecAction = NULL;
    //  QI for the executable task pointer.
    hr = pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);
    pAction->Release();
    if (FAILED(hr))
    {
        printf("\nQueryInterface call failed on IExecAction: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    //  Set the path of the executable to notepad.exe.
    hr = pExecAction->put_Path(_bstr_t(wstringExecutablePath.c_str()));
    pExecAction->Release();
    if (FAILED(hr))
    {
        printf("\nCannot add path for executable action: %x", hr);
        pRootFolder->Release();
        pTask->Release();
        CoUninitialize();
        return 1;
    }

    /*
    //  ------------------------------------------------------
    //  Securely get the user name and password. The task will
    //  be created to run with the credentials from the supplied 
    //  user name and password.
    CREDUI_INFO cui;
    TCHAR pszName[CREDUI_MAX_USERNAME_LENGTH];
    TCHAR pszPwd[CREDUI_MAX_PASSWORD_LENGTH];
    BOOL fSave;
    DWORD dwErr;

    cui.cbSize = sizeof(CREDUI_INFO);
    cui.hwndParent = NULL;
    //  Ensure that MessageText and CaptionText identify
    //  what credentials to use and which application requires them.
    cui.pszMessageText = TEXT("Account information for task registration:");
    cui.pszCaptionText = TEXT("Enter Account Information for Task Registration");
    cui.hbmBanner = NULL;
    fSave = FALSE;

    //  Create the UI asking for the credentials.
    dwErr = CredUIPromptForCredentials(
        &cui,                             //  CREDUI_INFO structure
        TEXT(""),                         //  Target for credentials
        NULL,                             //  Reserved
        0,                                //  Reason
        pszName,                          //  User name
        CREDUI_MAX_USERNAME_LENGTH,       //  Max number for user name
        pszPwd,                           //  Password
        CREDUI_MAX_PASSWORD_LENGTH,       //  Max number for password
        &fSave,                           //  State of save check box
        CREDUI_FLAGS_GENERIC_CREDENTIALS |  //  Flags
        CREDUI_FLAGS_ALWAYS_SHOW_UI |
        CREDUI_FLAGS_DO_NOT_PERSIST);

    if (dwErr)
    {
        std::cout << "Did not get credentials." << std::endl;
        CoUninitialize();
        return 1;
    }
    */

    //  ------------------------------------------------------
    //  Save the task in the root folder.
    IRegisteredTask* pRegisteredTask = NULL;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t(wszTaskName),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(),//_bstr_t(pszName)),
        _variant_t(),//_bstr_t(pszPwd)),
        TASK_LOGON_INTERACTIVE_TOKEN,
        _variant_t(L""),
        &pRegisteredTask);
    if (FAILED(hr))
    {
        printf("\nError saving the Task : %x", hr);
        pRootFolder->Release();
        pTask->Release();
        //SecureZeroMemory(pszName, sizeof(pszName));
        //SecureZeroMemory(pszPwd, sizeof(pszPwd));
        CoUninitialize();
        return 1;
    }

    printf("\n Success! Task succesfully registered. ");

    //  Clean up
    pRootFolder->Release();
    pTask->Release();
    pRegisteredTask->Release();
    //SecureZeroMemory(pszName, sizeof(pszName));
    //SecureZeroMemory(pszPwd, sizeof(pszPwd));
    CoUninitialize();
    return 0;
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
