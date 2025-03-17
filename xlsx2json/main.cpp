#include <iostream>
#include <random>
#include <string>
#include <iomanip>
#include <chrono>
#include <vector>
#include <fstream>
#include <mimalloc.h>
#include <xlnt/xlnt.hpp>
#include <json/json.h>
#include "imgui.h"
#include "imgui_impl_glfw.h"
#include "imgui_impl_opengl3.h"
#include <windows.h>
#include <GLFW/glfw3.h>
#include <glad/glad.h>
#include <oleidl.h>
#include <filesystem>

#define GLFW_EXPOSE_NATIVE_WIN32
#include <GLFW/glfw3native.h>

namespace fs = std::filesystem;

std::string logText = "";

void LoggerDump(const char* content)
{
    logText.insert(0, "\n");
    logText.insert(0, content);
    logText.insert(0, "\t");
    auto now = std::chrono::system_clock::now();
    auto duration = now.time_since_epoch();
    auto millis = std::chrono::duration_cast<std::chrono::milliseconds>(duration);
    std::time_t now_c = std::chrono::system_clock::to_time_t(now);
    std::tm* local_time = std::localtime(&now_c);
    std::ostringstream oss;
    oss << std::put_time(local_time, "%c");
    oss << "." << std::setfill('0') << std::setw(3) << millis.count() % 1000;
    logText.insert(0, oss.str());
}

// 将UTF-8字符串转换为宽字符（UTF-16）
std::wstring Utf8ToUtf16(const std::string& utf8_str) {
    int wide_size = MultiByteToWideChar(CP_UTF8, 0, &utf8_str[0], (int)utf8_str.size(), NULL, 0);
    std::vector<wchar_t> wstr(wide_size + 1, 0); // +1 for null terminator
    MultiByteToWideChar(CP_UTF8, 0, &utf8_str[0], -1, &wstr[0], wide_size);
    return std::wstring(&wstr[0]);
}

// 将宽字符（UTF-16）转换为GBK
std::string Utf16ToGbk(const std::wstring& utf16_str) {
    int gbk_size = WideCharToMultiByte(CP_ACP, 0, &utf16_str[0], (int)utf16_str.size(), NULL, 0, NULL, NULL);
    std::vector<char> gbk(gbk_size + 1, 0); // +1 for null terminator
    WideCharToMultiByte(CP_ACP, 0, &utf16_str[0], -1, &gbk[0], gbk_size, NULL, NULL);
    return std::string(&gbk[0]);
}

std::wstring CharToWchar(const std::string& mbstr) {
    const char* mbcstr = mbstr.c_str();
    int len = MultiByteToWideChar(CP_UTF8, 0, mbcstr, -1, nullptr, 0);
    if (len == 0) {
        std::cerr << "MultiByteToWideChar failed" << std::endl;
        return L"";
    }

    wchar_t* wcstr = new wchar_t[len];
    MultiByteToWideChar(CP_UTF8, 0, mbcstr, -1, wcstr, len);
    std::wstring result(wcstr);
    delete[] wcstr;
    return result;
}

std::string WcharToChar(const std::wstring& wstr) {
    const wchar_t* wccstr = wstr.c_str();
    int len = WideCharToMultiByte(CP_UTF8, 0, wccstr, -1, nullptr, 0, nullptr, nullptr);
    if (len == 0) {
        std::cerr << "WideCharToMultiByte failed" << std::endl;
        return "";
    }

    std::vector<char> mbstr(len);
    WideCharToMultiByte(CP_UTF8, 0, wccstr, -1, &mbstr[0], len, nullptr, nullptr);
    return std::string(mbstr.begin(), mbstr.end());
}

std::wstring changeFileExtension(const std::wstring& path, const std::wstring& newExtension) {
    // 获取文件的基本名和扩展名  
    std::wstring baseName = fs::path(path).filename().wstring();
    std::wstring extension = fs::path(path).extension().wstring();
    // 移除原始扩展名  
    size_t dotPos = baseName.rfind('.');
    if (dotPos != std::string::npos) {
        baseName.erase(dotPos);
    }
    // 添加新扩展名  
    return fs::path(path).parent_path().wstring() + L"\\" + baseName + newExtension;
}

void Xlsx2Json(std::wstring& srcPath, std::wstring& desPath)
{
    LoggerDump(WcharToChar(L" =================转换开始=================").c_str());
    xlnt::workbook wb;
    wb.load(srcPath);
    auto ws = wb.active_sheet();
    auto dim = ws.calculate_dimension();
    size_t max_row = dim.height();
    size_t max_column = dim.width();
    LoggerDump(WcharToChar(std::wstring(L"列数:") + std::to_wstring(max_column) + L"\t" + std::wstring(L"行数:") + std::to_wstring(max_row)).c_str());
    
    Json::Value arr;
    std::vector<std::string> keys;
    for (size_t row_index = 1; row_index <= max_row; ++row_index)
    {
        Json::Value tab;
        for (int32_t col_index = 1; col_index <= max_column; ++col_index)
        {
            xlnt::cell cell = ws.cell(col_index, row_index);
            if (row_index == 1) // 第一行读取键
            {
                keys.emplace_back(cell.to_string());
            }
            else
            {
                auto& key = keys[col_index - 1];
                if (cell.has_value())
                    tab[key] = cell.to_string();
                else
                    tab[key] = "";
            }
        }
        if (row_index == 1)
        {
            Json::Value keyJson;
            for (auto& key : keys)
            {
                keyJson.append(key);
            }
            Json::StreamWriterBuilder builder;
            builder.settings_["indentation"] = "";
            builder.settings_["emitUTF8"] = true;
            std::unique_ptr<Json::StreamWriter> writer(builder.newStreamWriter());
            std::ostringstream ossJson;
            writer->write(keyJson, &ossJson);
            LoggerDump(ossJson.str().c_str());
        }
        else
        {
            arr.append(tab);
        }
    }

    Json::StreamWriterBuilder builder;
    builder.settings_["indentation"] = "    ";
    builder.settings_["emitUTF8"] = true;
    std::unique_ptr<Json::StreamWriter> writer(builder.newStreamWriter());
    std::ostringstream ossJson;
    writer->write(arr, &ossJson);
    std::wstring utf16Str = Utf8ToUtf16(ossJson.str());
    std::string gbkStr = Utf16ToGbk(utf16Str);
    std::ofstream outPut(desPath.c_str());
    outPut << gbkStr << std::endl;
    LoggerDump(WcharToChar(std::wstring(L" 转换完成 ") + desPath).c_str());
}

void DrawConvertWindow()
{
    ImGui::SetNextWindowPos(ImVec2(0, 0), ImGuiCond_Always);
    ImGui::SetNextWindowSize(ImVec2(800, 400));

    ImGuiWindowFlags window_flags = 0;
    window_flags |= ImGuiWindowFlags_NoMove;
    window_flags |= ImGuiWindowFlags_NoBackground;
    ImGui::Begin(WcharToChar(std::wstring(L" 拖入xlsx文件生成json ")).c_str(), nullptr, window_flags);

    ImGui::SetWindowFontScale(0.5);
    ImGui::TextWrapped("%s", logText.c_str());
    ImGui::SetWindowFontScale(1);
    
    ImGui::End();
}

class DropManager : public IDropTarget
{
public:
    ULONG AddRef() { return 1; }
    ULONG Release() { return 0; }

    HRESULT QueryInterface(REFIID riid, void** ppvObject)
    {
        if (riid == IID_IDropTarget)
        {
            *ppvObject = this;
            return S_OK;
        }
        *ppvObject = NULL;
        return E_NOINTERFACE;
    };

    HRESULT DragEnter(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
    {
        // 处理拖入操作
        *pdwEffect &= DROPEFFECT_COPY;
        return S_OK;
    }

    HRESULT DragLeave()
    {
        return S_OK;
    }

    HRESULT DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
    {
        *pdwEffect &= DROPEFFECT_COPY;
        return S_OK;
    }

    HRESULT Drop(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect)
    {
        // 处理放置操作
        *pdwEffect &= DROPEFFECT_COPY;
        FORMATETC fmtetc = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        STGMEDIUM stgmedium;
        ZeroMemory(&stgmedium, sizeof(stgmedium));
        HRESULT hr = pDataObj->GetData(&fmtetc, &stgmedium);
        if (SUCCEEDED(hr))
        {
            HDROP hDrop = (HDROP)stgmedium.hGlobal;
            wchar_t xlsxPath[MAX_PATH] = {0};
            DragQueryFileW(hDrop, 0, xlsxPath, MAX_PATH);
            try { // 读取单元格的值
                std::wstring srcPath = xlsxPath;
                std::wstring desPath = changeFileExtension(xlsxPath, L".json");
                Xlsx2Json(srcPath, desPath);
            }
            catch (const std::exception& e) {
                LoggerDump((std::string("Error:") += e.what()).c_str());
            }
            DragFinish(hDrop);
            ReleaseStgMedium(&stgmedium);
        }
        return S_OK;
    }
};

int main(int argc, char* argv[])
{
    std::cout << "Tool Launch Success!" << std::endl;
    // 初始化 GLFW
    if (!glfwInit())
    {
        return -1;
    }
    ShowWindow(GetConsoleWindow(), SW_HIDE);
    //SetConsoleOutputCP(CP_UTF8);
    // 创建窗口和上下文
    GLFWwindow* window = glfwCreateWindow(800, 400, "xlsx2json create by ImGui", NULL, NULL);
    if (!window)
    {
        glfwTerminate();
        return -1;
    }
    glfwMakeContextCurrent(window);
    glfwSwapInterval(1);
    // 初始化 OpenGL3
    if (!gladLoadGLLoader((GLADloadproc)glfwGetProcAddress))
    {
        return -1;
    }
    // 初始化 Dear ImGui
    IMGUI_CHECKVERSION();
    ImGui::CreateContext();
    ImGui_ImplGlfw_InitForOpenGL(window, true);
    ImGui_ImplOpenGL3_Init("#version 130");

    ImGuiIO& io = ImGui::GetIO();
    ImFont* font = io.Fonts->AddFontFromFileTTF("C:\\Windows\\Fonts\\msyh.ttc", 28.0f, NULL, io.Fonts->GetGlyphRangesChineseSimplifiedCommon());
    IM_ASSERT(font != NULL);
    ImGuiStyle* style = &ImGui::GetStyle();
    style->Colors[ImGuiCol_Text] = ImVec4(0.0f, 1.0f, 0.0f, 1.0f);

    // 初始化拖文件相关
    OleInitialize(NULL);
    HWND hwnd = glfwGetWin32Window(window);
    DropManager dm;
    RegisterDragDrop(hwnd, &dm);

    // 主循环
    while (!glfwWindowShouldClose(window))
    {
        glfwPollEvents();

        // 开始 Dear ImGui 帧
        ImGui_ImplOpenGL3_NewFrame();
        ImGui_ImplGlfw_NewFrame();
        ImGui::NewFrame();
        
        DrawConvertWindow();

        // 渲染 Dear ImGui
        ImGui::Render();
        int display_w, display_h;
        glfwGetFramebufferSize(window, &display_w, &display_h);
        glViewport(0, 0, display_w, display_h);
        glClearColor(0.2f, 0.0f, 0.2f, 1.0f);
        glClear(GL_COLOR_BUFFER_BIT);
        ImGui_ImplOpenGL3_RenderDrawData(ImGui::GetDrawData());

        glfwSwapBuffers(window);
    }

    // 清理和销毁 Dear ImGui 上下文
    ImGui_ImplOpenGL3_Shutdown();
    ImGui_ImplGlfw_Shutdown();
    ImGui::DestroyContext();

    // 清理 GLFW 和 OpenGL3
    glfwDestroyWindow(window);
    glfwTerminate();
    
    // 清理拖文件相关
    RevokeDragDrop(hwnd);
    OleUninitialize();

    return 0;
}
