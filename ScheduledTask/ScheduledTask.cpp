// ScheduledTask.cpp : This file contains the 'main' function. Program execution begins and ends there.
//
#include <string>
#include <sstream>
#include <iostream>
#include <iomanip>

#define _WIN32_DCOM

#include <windows.h>
#include <atlbase.h>
#include <Shlobj.h>
#include <comdef.h>
#include <taskschd.h>
#pragma comment(lib, "taskschd.lib")

#define THROW_IF_FAILED(hr) if(FAILED(hr)){ throw HrException(hr); }
#define THROW_IF_FAILED_MSG(hr, msg) if(FAILED(hr)){ throw HrException(hr, msg); }

class HrException : public std::exception
{
private:
    HRESULT _hr;
    std::wstring _message;

    static std::wstring GetErrorMessage(HRESULT hr)
    {
        _com_error err(hr);
        std::wostringstream msg;
        msg << err.ErrorMessage() << " (0x" << std::hex << hr << ")";
        return msg.str();
    }

    static std::wstring GetErrorMessage(HRESULT hr, const std::wstring& message)
    {
        std::wostringstream msg;
        msg << message << ": " << GetErrorMessage(hr);
        return msg.str();
    }

public:
    HrException(HRESULT hr) : _hr(hr) {}
    HrException(HRESULT hr, const std::wstring& message) : _hr(hr), _message(message) {}

    virtual std::wstring message() const
    {
        return _message.empty() ? GetErrorMessage(_hr) : GetErrorMessage(_hr, _message);
    }
};

void CreateTask(int argc, wchar_t** argv);
HRESULT ParseCommandLineArgs(int argc, wchar_t** argv, bool& useSystem, bool& deleteTask, std::wstring& cmd, std::wstring& arguments);
std::wstring GetDefaultCommand();

int wmain(int argc, wchar_t** argv)
{
    try
    {
        //  Initialize COM.
        HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
        THROW_IF_FAILED_MSG(hr, L"CoInitializeEx failed");

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
        THROW_IF_FAILED_MSG(hr, L"CoInitializeSecurity failed");

        CreateTask(argc, argv);
    }
    catch (const HrException& ex)
    {
        std::wcerr << ex.message() << std::endl;
        CoUninitialize();
        return 1;
    }
    catch (...)
    {
        std::wcerr << L"Unexpected error!" << std::endl;
        CoUninitialize();
        return 1;
    }

    CoUninitialize();

    return 0;
}

void CreateTask(int argc, wchar_t** argv)
{
    std::wstring cmd, arguments;
    bool useSystem = false, deleteTask = false;
    HRESULT hr = ParseCommandLineArgs(argc, argv, useSystem, deleteTask, cmd, arguments);
    THROW_IF_FAILED(hr);

    //  ------------------------------------------------------
    //  Create a name for the task.
    LPCWSTR wszTaskName = L"netspi-native-exe-task";

    //  Create an instance of the Task Service. 
    CComPtr<ITaskService> pService;
    hr = CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pService));
    THROW_IF_FAILED_MSG(hr, L"Failed to create an instance of ITaskService");

    //  Connect to the task service.
    hr = pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
    THROW_IF_FAILED_MSG(hr, L"ITaskService::Connect failed");

    //  Get the pointer to the root task folder.  
    //  This folder will hold the new task that is registered.
    CComPtr<ITaskFolder> pRootFolder;
    hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
    THROW_IF_FAILED_MSG(hr, L"Cannot get Root Folder pointer");

    //  If the same task exists, remove it.
    hr = pRootFolder->DeleteTask(_bstr_t(wszTaskName), 0);

    if (deleteTask)
    {
        THROW_IF_FAILED_MSG(hr, L"Failed to delete task definition");

        std::wcout << L"Success! Task successfully deleted." << std::endl;
        return;
    }

    //  Create the task builder object to create the task.
    CComPtr<ITaskDefinition> pTask;
    hr = pService->NewTask(0, &pTask);
    THROW_IF_FAILED_MSG(hr, L"Failed to create a task definition");

    //  ------------------------------------------------------
    //  Get the registration info for setting the identification.
    CComPtr<IRegistrationInfo> pRegInfo;
    hr = pTask->get_RegistrationInfo(&pRegInfo);
    THROW_IF_FAILED_MSG(hr, L"Cannot get identification pointer");

    hr = pRegInfo->put_Author(SysAllocString(L"NetSPI"));
    THROW_IF_FAILED_MSG(hr, L"Cannot put identification info");

    ////////////////////////////////////////////////////////////////////////////////
    ///
    ////////////////////////////////////////////////////////////////////////////////

    if (useSystem)
    {
        //  Create the principal for the task
        CComPtr<IPrincipal> pPrincipal;
        hr = pTask->get_Principal(&pPrincipal);
        THROW_IF_FAILED_MSG(hr, L"Cannot get principal pointer");

        //  Set up principal information: 
        hr = pPrincipal->put_Id(_bstr_t(L"SYSTEM"));
        if (FAILED(hr))
        {
            std::wcerr << L"Cannot put the principal ID: " << std::hex << hr << std::endl;
        }

        hr = pPrincipal->put_LogonType(TASK_LOGON_INTERACTIVE_TOKEN);
        if (FAILED(hr))
        {
            std::wcerr << L"Cannot put principal logon type: " << std::hex << hr << std::endl;
        }

        //  Run the task with the least privileges (LUA) 
        hr = pPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);
        THROW_IF_FAILED_MSG(hr, L"Cannot put principal run level");
    }


    // Configure Settings
    CComPtr<ITaskSettings> pSettings;
    hr = pTask->get_Settings(&pSettings);
    THROW_IF_FAILED_MSG(hr, L"Cannot get settings pointer");

    //  Set setting values for the task.
    hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
    THROW_IF_FAILED_MSG(hr, L"Cannot put setting info");


    // Configure Trigger
    CComPtr<ITriggerCollection> pTriggerCollection;
    hr = pTask->get_Triggers(&pTriggerCollection);
    THROW_IF_FAILED_MSG(hr, L"Cannot get trigger collection");

    CComPtr<ITrigger> pTrigger;
    hr = pTriggerCollection->Create(TASK_TRIGGER_REGISTRATION, &pTrigger);
    THROW_IF_FAILED_MSG(hr, L"Cannot create a registration trigger");

    CComPtr<IRegistrationTrigger> pRegistrationTrigger;
    hr = pTrigger->QueryInterface(IID_PPV_ARGS(&pRegistrationTrigger));
    THROW_IF_FAILED_MSG(hr, L"QueryInterface call failed on IRegistrationTrigger");

    hr = pRegistrationTrigger->put_Id(_bstr_t(L"OneTimeTrigger"));
    if (FAILED(hr))
    {
        std::wcerr << L"Cannot put trigger ID: " << std::hex << hr << std::endl;
    }

    //  Define the delay for the registration trigger.
    hr = pRegistrationTrigger->put_Delay(_bstr_t(L"PT30S"));
    THROW_IF_FAILED_MSG(hr, L"Cannot put registration trigger delay");


    // Configure Action
    CComPtr<IActionCollection> pActionCollection;
    hr = pTask->get_Actions(&pActionCollection);
    THROW_IF_FAILED_MSG(hr, L"Cannot get trigger collection");

    //  Create the action, specifying that it is an executable action.
    CComPtr<IAction> pAction;
    hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
    THROW_IF_FAILED_MSG(hr, L"Cannot create the action");

    //  QI for the executable task pointer.
    CComPtr<IExecAction> pExecAction;
    hr = pAction->QueryInterface(IID_PPV_ARGS(&pExecAction));
    THROW_IF_FAILED_MSG(hr, L"QueryInterface call failed on IExecAction");

    //  Set the path of the executable.
    if (cmd.empty())
    {
        cmd = GetDefaultCommand();
    }

    hr = pExecAction->put_Path(_bstr_t(cmd.c_str()));
    THROW_IF_FAILED_MSG(hr, L"Cannot add path for executable action");

    if (!arguments.empty())
    {
        hr = pExecAction->put_Arguments(_bstr_t(arguments.c_str()));
        THROW_IF_FAILED_MSG(hr, L"Cannot add arguments for executable action");
    }

    //  Save the task in the root folder.
    CComPtr<IRegisteredTask> pRegisteredTask;
    hr = pRootFolder->RegisterTaskDefinition(
        _bstr_t(wszTaskName),
        pTask,
        TASK_CREATE_OR_UPDATE,
        _variant_t(),
        _variant_t(),
        TASK_LOGON_INTERACTIVE_TOKEN,
        _variant_t(L""),
        &pRegisteredTask);
    THROW_IF_FAILED_MSG(hr, L"Error saving the Task");

    std::wcout << L"Success! Task successfully registered." << std::endl;
}

HRESULT ParseCommandLineArgs(int argc, wchar_t** argv, bool& useSystem, bool& deleteTask, std::wstring& cmd, std::wstring& arguments)
{
    for (int i = 1; i < argc; i++)
    {
        std::wstring arg(argv[i]);
        if (arg == L"/s")
        {
            useSystem = true;
        }
        else if (arg == L"/d")
        {
            deleteTask = true;
        }
        else if (arg == L"/c")
        {
            if (++i < argc)
            {
                cmd = std::wstring(argv[i]);
            }
            else
            {
                return E_INVALIDARG;
            }
        }
        else if (arg == L"/a")
        {
            if (++i < argc)
            {
                arguments = std::wstring(argv[i]);
            }
            else
            {
                return E_INVALIDARG;
            }
        }
    }

    return S_OK;
}

std::wstring GetDefaultCommand()
{
    PWSTR pSystemPath = nullptr;
    HRESULT hr = SHGetKnownFolderPath(FOLDERID_System, 0, NULL, &pSystemPath);
    THROW_IF_FAILED_MSG(hr, L"");

    if (FAILED(hr))
    {
        std::wcerr << L"Unable to get system directory: " << std::hex << hr << std::endl;
        return std::wstring();
    }

    std::wstring wstringExecutablePath(pSystemPath);
    wstringExecutablePath += std::wstring(L"\\NOTEPAD.EXE");

    CoTaskMemFree(pSystemPath);

    return wstringExecutablePath;
}
