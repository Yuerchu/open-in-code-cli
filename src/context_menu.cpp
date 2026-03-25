#define WIN32_LEAN_AND_MEAN
#define UNICODE
#define _UNICODE

#include <windows.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <shobjidl_core.h>
#include <wrl/module.h>
#include <wrl/implements.h>
#include <wrl/client.h>
#include <string>
#include <vector>

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "runtimeobject.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "ole32.lib")

using Microsoft::WRL::ClassicCom;
using Microsoft::WRL::ComPtr;
using Microsoft::WRL::InhibitRoOriginateError;
using Microsoft::WRL::Make;
using Microsoft::WRL::Module;
using Microsoft::WRL::ModuleType;
using Microsoft::WRL::RuntimeClass;
using Microsoft::WRL::RuntimeClassFlags;

static HMODULE g_hModule = nullptr;

extern "C" BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        g_hModule = instance;
        DisableThreadLibraryCalls(instance);
    }
    return TRUE;
}

// ============================================================
// Helpers
// ============================================================
namespace {

std::wstring GetDllDirectory() {
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(g_hModule, path, MAX_PATH);
    std::wstring dir(path);
    auto pos = dir.find_last_of(L'\\');
    if (pos != std::wstring::npos) dir.resize(pos);
    return dir;
}

std::wstring GetIconPath(const wchar_t* name) {
    return GetDllDirectory() + L"\\" + name;
}

std::wstring QuoteArg(const std::wstring& arg) {
    if (arg.find_first_of(L" \\\"") == std::wstring::npos)
        return arg;
    std::wstring out;
    out.push_back(L'"');
    for (size_t i = 0; i < arg.size(); ++i) {
        if (arg[i] == L'\\') {
            size_t start = i, end = start + 1;
            for (; end < arg.size() && arg[end] == L'\\'; ++end) {}
            size_t n = end - start;
            if (end == arg.size() || arg[end] == L'"') n *= 2;
            for (size_t j = 0; j < n; ++j) out.push_back(L'\\');
            i = end - 1;
        } else if (arg[i] == L'"') {
            out.push_back(L'\\');
            out.push_back(L'"');
        } else {
            out.push_back(arg[i]);
        }
    }
    out.push_back(L'"');
    return out;
}

HRESULT LaunchWT(IShellItemArray* items, const wchar_t* argsTemplate) {
    if (!items) return S_OK;
    DWORD count = 0;
    if (FAILED(items->GetCount(&count))) return S_OK;
    for (DWORD i = 0; i < count; ++i) {
        ComPtr<IShellItem> item;
        if (SUCCEEDED(items->GetItemAt(i, &item))) {
            LPWSTR filePath = nullptr;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &filePath))) {
                std::wstring args(argsTemplate);
                // Replace {path} placeholder
                std::wstring quoted = QuoteArg(filePath);
                size_t pos;
                while ((pos = args.find(L"{path}")) != std::wstring::npos)
                    args.replace(pos, 6, quoted);
                while ((pos = args.find(L"{rawpath}")) != std::wstring::npos)
                    args.replace(pos, 9, filePath);
                ShellExecuteW(nullptr, L"open", L"wt.exe",
                              args.c_str(), nullptr, SW_SHOWNORMAL);
                CoTaskMemFree(filePath);
            }
        }
    }
    return S_OK;
}

} // namespace

// ============================================================
// Child Commands (not COM-registered)
// ============================================================

// Claude Code (新对话) — claude
class ClaudeNewCommand final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>,
                          IExplorerCommand> {
 public:
  IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name) {
      return SHStrDupW(L"Claude Code (\x65B0\x5BF9\x8BDD)", name);
  }
  IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* icon) {
      return SHStrDupW(GetIconPath(L"claude.ico").c_str(), icon);
  }
  IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* t) { *t = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP GetCanonicalName(GUID* g) { *g = GUID_NULL; return S_OK; }
  IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* s) { *s = ECS_ENABLED; return S_OK; }
  IFACEMETHODIMP GetFlags(EXPCMDFLAGS* f) { *f = ECF_DEFAULT; return S_OK; }
  IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** e) { *e = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP Invoke(IShellItemArray* items, IBindCtx*) {
      return LaunchWT(items, L"-d {path} -- claude");
  }
};

// Claude Code (继续对话) — claude -c
class ClaudeContinueCommand final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>,
                          IExplorerCommand> {
 public:
  IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name) {
      return SHStrDupW(L"Claude Code (\x7EE7\x7EED\x5BF9\x8BDD)", name);
  }
  IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* icon) {
      return SHStrDupW(GetIconPath(L"claude.ico").c_str(), icon);
  }
  IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* t) { *t = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP GetCanonicalName(GUID* g) { *g = GUID_NULL; return S_OK; }
  IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* s) { *s = ECS_ENABLED; return S_OK; }
  IFACEMETHODIMP GetFlags(EXPCMDFLAGS* f) { *f = ECF_DEFAULT; return S_OK; }
  IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** e) { *e = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP Invoke(IShellItemArray* items, IBindCtx*) {
      return LaunchWT(items, L"-d {path} -- claude -c");
  }
};

// Claude Code (从历史记录继续) — claude --resume
class ClaudeResumeCommand final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>,
                          IExplorerCommand> {
 public:
  IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name) {
      return SHStrDupW(L"Claude Code (\x4ECE\x5386\x53F2\x8BB0\x5F55\x7EE7\x7EED)", name);
  }
  IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* icon) {
      return SHStrDupW(GetIconPath(L"claude.ico").c_str(), icon);
  }
  IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* t) { *t = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP GetCanonicalName(GUID* g) { *g = GUID_NULL; return S_OK; }
  IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* s) { *s = ECS_ENABLED; return S_OK; }
  IFACEMETHODIMP GetFlags(EXPCMDFLAGS* f) { *f = ECF_DEFAULT; return S_OK; }
  IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** e) { *e = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP Invoke(IShellItemArray* items, IBindCtx*) {
      return LaunchWT(items, L"-d {path} -- claude --resume");
  }
};

// Separator
class SeparatorCommand final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>,
                          IExplorerCommand> {
 public:
  IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name) { *name = nullptr; return S_OK; }
  IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* i) { *i = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* t) { *t = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP GetCanonicalName(GUID* g) { *g = GUID_NULL; return S_OK; }
  IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* s) { *s = ECS_ENABLED; return S_OK; }
  IFACEMETHODIMP GetFlags(EXPCMDFLAGS* f) { *f = ECF_ISSEPARATOR; return S_OK; }
  IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** e) { *e = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP Invoke(IShellItemArray*, IBindCtx*) { return S_OK; }
};

// Codex (WSL)
class CodexCommand final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>,
                          IExplorerCommand> {
 public:
  IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name) {
      return SHStrDupW(L"Codex", name);
  }
  IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* icon) {
      return SHStrDupW(GetIconPath(L"codex.ico").c_str(), icon);
  }
  IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* t) { *t = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP GetCanonicalName(GUID* g) { *g = GUID_NULL; return S_OK; }
  IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* s) { *s = ECS_ENABLED; return S_OK; }
  IFACEMETHODIMP GetFlags(EXPCMDFLAGS* f) { *f = ECF_DEFAULT; return S_OK; }
  IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** e) { *e = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP Invoke(IShellItemArray* items, IBindCtx*) {
      return LaunchWT(items, L"-- wsl --cd {path} bash -lic codex");
  }
};

// ============================================================
// Sub-command enumerator
// ============================================================
class SubCommandEnumerator final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>,
                          IEnumExplorerCommand> {
    std::vector<ComPtr<IExplorerCommand>> m_commands;
    DWORD m_current = 0;

 public:
    SubCommandEnumerator() {
        auto add = [this](auto obj) {
            ComPtr<IExplorerCommand> cmd;
            obj->QueryInterface(IID_PPV_ARGS(&cmd));
            m_commands.push_back(std::move(cmd));
        };
        add(Make<ClaudeNewCommand>());
        add(Make<ClaudeContinueCommand>());
        add(Make<ClaudeResumeCommand>());
        add(Make<SeparatorCommand>());
        add(Make<CodexCommand>());
    }

    IFACEMETHODIMP Next(ULONG celt, IExplorerCommand** rgelt, ULONG* pceltFetched) {
        ULONG fetched = 0;
        while (fetched < celt && m_current < m_commands.size()) {
            m_commands[m_current].CopyTo(&rgelt[fetched]);
            m_current++;
            fetched++;
        }
        if (pceltFetched) *pceltFetched = fetched;
        return (fetched == celt) ? S_OK : S_FALSE;
    }

    IFACEMETHODIMP Skip(ULONG celt) {
        m_current = min(m_current + celt, (DWORD)m_commands.size());
        return (m_current < m_commands.size()) ? S_OK : S_FALSE;
    }

    IFACEMETHODIMP Reset() {
        m_current = 0;
        return S_OK;
    }

    IFACEMETHODIMP Clone(IEnumExplorerCommand** ppenum) {
        *ppenum = nullptr;
        return E_NOTIMPL;
    }
};

// ============================================================
// Parent Command — {E8A1D5B2-7C3F-4A9E-B6D4-1F2E3A4B5C6D}
// ============================================================
class __declspec(uuid("E8A1D5B2-7C3F-4A9E-B6D4-1F2E3A4B5C6D"))
    OpenInCodeCLICommand final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>,
                          IExplorerCommand> {
 public:
  IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name) {
      return SHStrDupW(L"Open in Code CLI", name);
  }
  IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* icon) {
      // Use Windows Terminal icon
      return SHStrDupW(L"wt.exe", icon);
  }
  IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* t) { *t = nullptr; return E_NOTIMPL; }
  IFACEMETHODIMP GetCanonicalName(GUID* g) { *g = GUID_NULL; return S_OK; }
  IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* s) { *s = ECS_ENABLED; return S_OK; }
  IFACEMETHODIMP GetFlags(EXPCMDFLAGS* f) { *f = ECF_HASSUBCOMMANDS; return S_OK; }
  IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** e) {
      auto enumerator = Make<SubCommandEnumerator>();
      *e = enumerator.Detach();
      return S_OK;
  }
  IFACEMETHODIMP Invoke(IShellItemArray*, IBindCtx*) { return S_OK; }
};

// ============================================================
// COM registration — only the parent command
// ============================================================
CoCreatableClass(OpenInCodeCLICommand)
CoCreatableClassWrlCreatorMapInclude(OpenInCodeCLICommand)

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    if (ppv == nullptr) return E_POINTER;
    *ppv = nullptr;
    return Module<ModuleType::InProc>::GetModule().GetClassObject(rclsid, riid, ppv);
}

STDAPI DllCanUnloadNow(void) {
    return Module<ModuleType::InProc>::GetModule().GetObjectCount() == 0
               ? S_OK : S_FALSE;
}

STDAPI DllGetActivationFactory(HSTRING activatableClassId,
                               IActivationFactory** factory) {
    return Module<ModuleType::InProc>::GetModule()
        .GetActivationFactory(activatableClassId, factory);
}
