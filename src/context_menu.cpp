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

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "runtimeobject.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "ole32.lib")

using Microsoft::WRL::ClassicCom;
using Microsoft::WRL::ComPtr;
using Microsoft::WRL::InhibitRoOriginateError;
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

namespace {

std::wstring GetDllDirectory() {
    wchar_t path[MAX_PATH];
    GetModuleFileNameW(g_hModule, path, MAX_PATH);
    std::wstring dir(path);
    auto pos = dir.find_last_of(L'\\');
    if (pos != std::wstring::npos) dir.resize(pos);
    return dir;
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

} // namespace

// ============================================================
// Claude Code - {E8A1D5B2-7C3F-4A9E-B6D4-1F2E3A4B5C6D}
// ============================================================
class __declspec(uuid("E8A1D5B2-7C3F-4A9E-B6D4-1F2E3A4B5C6D"))
    ClaudeCodeCommand final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>,
                          IExplorerCommand> {
 public:
  IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name) {
      // "使用 Claude Code 打开"
      return SHStrDupW(L"\x4F7F\x7528 Claude Code \x6253\x5F00", name);
  }
  IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* icon) {
      std::wstring path = GetDllDirectory() + L"\\claude.ico";
      return SHStrDupW(path.c_str(), icon);
  }
  IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* tip) {
      *tip = nullptr; return E_NOTIMPL;
  }
  IFACEMETHODIMP GetCanonicalName(GUID* guid) {
      *guid = GUID_NULL; return S_OK;
  }
  IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* state) {
      *state = ECS_ENABLED; return S_OK;
  }
  IFACEMETHODIMP GetFlags(EXPCMDFLAGS* flags) {
      *flags = ECF_DEFAULT; return S_OK;
  }
  IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** e) {
      *e = nullptr; return E_NOTIMPL;
  }
  IFACEMETHODIMP Invoke(IShellItemArray* items, IBindCtx*) {
      if (!items) return S_OK;
      DWORD count = 0;
      if (FAILED(items->GetCount(&count))) return S_OK;
      for (DWORD i = 0; i < count; ++i) {
          ComPtr<IShellItem> item;
          if (SUCCEEDED(items->GetItemAt(i, &item))) {
              LPWSTR filePath = nullptr;
              if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &filePath))) {
                  std::wstring args = L"-d " + QuoteArg(filePath) + L" -- claude";
                  ShellExecuteW(nullptr, L"open", L"wt.exe",
                                args.c_str(), nullptr, SW_SHOWNORMAL);
                  CoTaskMemFree(filePath);
              }
          }
      }
      return S_OK;
  }
};

// ============================================================
// Codex (WSL) - {F9B2E6C3-8D4A-4BAF-C7E5-2A3F4B5C6D7E}
// ============================================================
class __declspec(uuid("F9B2E6C3-8D4A-4BAF-C7E5-2A3F4B5C6D7E"))
    CodexCommand final
    : public RuntimeClass<RuntimeClassFlags<ClassicCom | InhibitRoOriginateError>,
                          IExplorerCommand> {
 public:
  IFACEMETHODIMP GetTitle(IShellItemArray*, PWSTR* name) {
      // "使用 Codex 打开"
      return SHStrDupW(L"\x4F7F\x7528 Codex \x6253\x5F00", name);
  }
  IFACEMETHODIMP GetIcon(IShellItemArray*, PWSTR* icon) {
      std::wstring path = GetDllDirectory() + L"\\codex.ico";
      return SHStrDupW(path.c_str(), icon);
  }
  IFACEMETHODIMP GetToolTip(IShellItemArray*, PWSTR* tip) {
      *tip = nullptr; return E_NOTIMPL;
  }
  IFACEMETHODIMP GetCanonicalName(GUID* guid) {
      *guid = GUID_NULL; return S_OK;
  }
  IFACEMETHODIMP GetState(IShellItemArray*, BOOL, EXPCMDSTATE* state) {
      *state = ECS_ENABLED; return S_OK;
  }
  IFACEMETHODIMP GetFlags(EXPCMDFLAGS* flags) {
      *flags = ECF_DEFAULT; return S_OK;
  }
  IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** e) {
      *e = nullptr; return E_NOTIMPL;
  }
  IFACEMETHODIMP Invoke(IShellItemArray* items, IBindCtx*) {
      if (!items) return S_OK;
      DWORD count = 0;
      if (FAILED(items->GetCount(&count))) return S_OK;
      for (DWORD i = 0; i < count; ++i) {
          ComPtr<IShellItem> item;
          if (SUCCEEDED(items->GetItemAt(i, &item))) {
              LPWSTR filePath = nullptr;
              if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &filePath))) {
                  // Use wsl --cd which auto-converts Windows paths
                  std::wstring args = L"-- wsl --cd " + QuoteArg(filePath)
                                    + L" bash -lic codex";
                  ShellExecuteW(nullptr, L"open", L"wt.exe",
                                args.c_str(), nullptr, SW_SHOWNORMAL);
                  CoTaskMemFree(filePath);
              }
          }
      }
      return S_OK;
  }
};

// COM registration
CoCreatableClass(ClaudeCodeCommand)
CoCreatableClass(CodexCommand)
CoCreatableClassWrlCreatorMapInclude(ClaudeCodeCommand)
CoCreatableClassWrlCreatorMapInclude(CodexCommand)

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
