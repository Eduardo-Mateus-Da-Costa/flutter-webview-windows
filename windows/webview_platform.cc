#include "webview_platform.h"
#include <cstdio>
#include <DispatcherQueue.h>
#include <shlobj.h>
#include <windows.graphics.capture.h>
#include <filesystem>
#include <iostream>

extern "C" HRESULT WINAPI CreateDispatcherQueueController(
        DispatcherQueueOptions options,
ABI::Windows::System::IDispatcherQueueController** dispatcherQueueController);
#pragma comment(lib, "WindowsApp.lib")

HRESULT CreateDispatcherQueueControllerOnSTA(
        DispatcherQueueOptions options,
        ABI::Windows::System::IDispatcherQueueController** controller) {

    HANDLE hEvent = CreateEvent(nullptr, FALSE, FALSE, nullptr);
    HRESULT hr = E_FAIL;

    std::thread staThread([&]() {
        CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        hr = CreateDispatcherQueueController(options, controller);
        SetEvent(hEvent);
        CoUninitialize();
    });

    staThread.detach();
    WaitForSingleObject(hEvent, INFINITE);
    CloseHandle(hEvent);

    return hr;
}


WebviewPlatform::WebviewPlatform()
        : rohelper_(std::make_unique<rx::RoHelper>(RO_INIT_SINGLETHREADED)) {

    if (rohelper_->WinRtAvailable()) {
        DispatcherQueueOptions options{
                sizeof(DispatcherQueueOptions),
                DQTYPE_THREAD_CURRENT,
                DQTAT_COM_STA
        };

        // 1. Tenta criar DispatcherQueue no thread atual
        HRESULT hr = rohelper_->CreateDispatcherQueueController(
                options, dispatcher_queue_controller_.put());

        // 2. Se falhar por erro de threading, tenta STA dedicado
        if (hr == 0x8001010E /* RPC_E_WRONG_THREAD */) {
            printf("CreateDispatcherQueueController falhou por thread errada, tentando STA dedicado...\n");
            fflush(stdout);

            hr = CreateDispatcherQueueControllerOnSTA(
                    options, dispatcher_queue_controller_.put());
        }

        // 3. Se ainda falhar, tenta fallback para RO_INIT_MULTITHREADED
        if (FAILED(hr)) {
            printf("Tentando fallback com RO_INIT_MULTITHREADED...\n");
            fflush(stdout);

            rohelper_ = std::make_unique<rx::RoHelper>(RO_INIT_MULTITHREADED);
            hr = rohelper_->CreateDispatcherQueueController(
                    options, dispatcher_queue_controller_.put());
        }

        // 4. Se continuar falhando, aborta
        if (FAILED(hr)) {
            printf("CreateDispatcherQueueController failed. HRESULT: 0x%08lX\n", hr);
            fflush(stdout);
            return;
        }

        // Verifica suporte ao GraphicsCapture
        if (!IsGraphicsCaptureSessionSupported()) {
            printf("Windows::Graphics::Capture::GraphicsCaptureSession is not supported.\n");
            fflush(stdout);
            return;
        }

        graphics_context_ = std::make_unique<GraphicsContext>(rohelper_.get());
        valid_ = graphics_context_->IsValid();
    }
}

bool WebviewPlatform::IsGraphicsCaptureSessionSupported() {
  HSTRING className;
  HSTRING_HEADER classNameHeader;

  if (FAILED(rohelper_->GetStringReference(
          RuntimeClass_Windows_Graphics_Capture_GraphicsCaptureSession,
          &className, &classNameHeader))) {
      printf("Failed to get string reference for Windows::Graphics::Capture::GraphicsCaptureSession.\n");
        fflush(stdout);
    return false;
  }

  ABI::Windows::Graphics::Capture::IGraphicsCaptureSessionStatics*
      capture_session_statics;
  if (FAILED(rohelper_->GetActivationFactory(
          className,
          __uuidof(
              ABI::Windows::Graphics::Capture::IGraphicsCaptureSessionStatics),
          (void**)&capture_session_statics))) {
        printf("Failed to get activation factory for Windows::Graphics::Capture::GraphicsCaptureSession.\n");
        fflush(stdout);
    return false;
  }

  boolean is_supported = false;
  if (FAILED(capture_session_statics->IsSupported(&is_supported))) {
        printf("Failed to check if Windows::Graphics::Capture::GraphicsCaptureSession is supported.\n");
        fflush(stdout);
    return false;
  }

  return !!is_supported;
}

std::optional<std::wstring> WebviewPlatform::GetDefaultDataDirectory() {
  PWSTR path_tmp;
  if (!SUCCEEDED(
          SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &path_tmp))) {
    return std::nullopt;
  }
  auto path = std::filesystem::path(path_tmp);
  CoTaskMemFree(path_tmp);

  wchar_t filename[MAX_PATH];
  GetModuleFileName(nullptr, filename, MAX_PATH);
  path /= "flutter_webview_windows";
  path /= std::filesystem::path(filename).stem();

  return path.wstring();
}
