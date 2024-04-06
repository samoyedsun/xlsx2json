#include <iostream>
#include <random>
#include <string>
#include <iomanip>
#include <chrono>
#include <vector>
#include <fstream>
#include <mimalloc.h>
#include <xlnt/xlnt.hpp>
#include <nlohmann/json.hpp>
#include "imgui.h"
#include "imgui_impl_glfw.h"
#include "imgui_impl_opengl3.h"
#include <GLFW/glfw3.h>
#include <glad/glad.h>
#include <windows.h>

using ordered_json = nlohmann::ordered_json;

static std::string logText = "";
static char xlsxPath[1024] = "E:\\Workspaces\\ProjectMir9\\Mir9_Client\\Database\\";
static char jsonPath[1024] = "E:\\Workspaces\\ProjectMir9\\Mir9_Client\\Database\\";
static char xlsxFile[1024] = "mydb_magic_tbl.xlsx";
static char jsonFile[1024] = "mydb_magic_tbl.json";

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

std::wstring CharToWchar(const char* mbstr) {
    if (mbstr == nullptr) return L"";

    int len = MultiByteToWideChar(CP_UTF8, 0, mbstr, -1, nullptr, 0);
    if (len == 0) {
        std::cerr << "MultiByteToWideChar failed" << std::endl;
        return L"";
    }

    wchar_t* wcstr = new wchar_t[len];
    MultiByteToWideChar(CP_UTF8, 0, mbstr, -1, wcstr, len);
    std::wstring result(wcstr);
    delete[] wcstr;
    return result;
}

std::string WcharToChar(const wchar_t* wcstr) {
    if (wcstr == nullptr) return "";

    int len = WideCharToMultiByte(CP_UTF8, 0, wcstr, -1, nullptr, 0, nullptr, nullptr);
    if (len == 0) {
        std::cerr << "WideCharToMultiByte failed" << std::endl;
        return "";
    }

    std::vector<char> mbstr(len);
    WideCharToMultiByte(CP_UTF8, 0, wcstr, -1, &mbstr[0], len, nullptr, nullptr);
    return std::string(mbstr.begin(), mbstr.end());
}

void Xlsx2Json()
{
    std::string xlslRealPath = xlsxPath;
    xlslRealPath += xlsxFile;
    LoggerDump((std::string(" 转换启动 ") += xlslRealPath).c_str());
    xlnt::workbook wb;
    wb.load(xlslRealPath);
    auto ws = wb.active_sheet();
    auto dim = ws.calculate_dimension();
    std::string logContent = std::string("列数:") + std::to_string(dim.width());
    logContent += "\t" + std::string("行数:") + std::to_string(dim.height());
    LoggerDump(logContent.c_str());
    
    ordered_json arr;
    std::vector<std::string> keys;
    int32_t line = 0;
    int32_t column = 0;
    for (const auto& row : ws)
    {
        ordered_json tab;
        for (const auto& cell : row)
        {
            if (line == 0) // 第一行读取键
            {
                keys.emplace_back(cell.to_string());
            }
            else
            {
                auto& key = keys[column];
                tab[key] = cell.to_string();
            }
            column++;
        }
        if (line == 0)
        {
            ordered_json keyJson;
            for (auto& key : keys)
            {
                keyJson.emplace_back(key);
            }
            LoggerDump(keyJson.dump().c_str());
        }
        else
        {
            arr.push_back(tab);
        }
        column = 0;
        line++;
    }

    std::string jsonRealPath = jsonPath;
    jsonRealPath += jsonFile;
    std::ofstream outPut(CharToWchar(jsonRealPath.c_str()));
    outPut << std::setw(4) << arr << std::endl;
    LoggerDump((std::string(" 转换完成 ") += jsonRealPath).c_str());
}

void SetCorner(int corner = 0)
{
    if (corner != -1)
    {
        const float PAD = 15.0f;
        const ImGuiViewport* viewport = ImGui::GetMainViewport();
        ImVec2 work_pos = viewport->WorkPos;
        ImVec2 work_size = viewport->WorkSize;
        ImVec2 window_pos, window_pos_pivot;
        window_pos.x = (corner & 1) ? (work_pos.x + work_size.x - PAD) : (work_pos.x + PAD);
        window_pos.y = (corner & 2) ? (work_pos.y + work_size.y - PAD) : (work_pos.y + PAD);
        window_pos_pivot.x = (corner & 1) ? 1.0f : 0.0f;
        window_pos_pivot.y = (corner & 2) ? 1.0f : 0.0f;
        ImGui::SetNextWindowPos(window_pos, ImGuiCond_Always, window_pos_pivot);
    }
}

void DrawConvertWindow()
{
    SetCorner(0);
    ImGui::SetNextWindowSize(ImVec2(700, 500));

    ImGuiWindowFlags window_flags = 0;
    window_flags |= ImGuiWindowFlags_NoDecoration;
    window_flags |= ImGuiWindowFlags_NoMove;
    //window_flags |= ImGuiWindowFlags_NoSavedSettings;
    //window_flags |= ImGuiWindowFlags_NoFocusOnAppearing;
    //window_flags |= ImGuiWindowFlags_NoBackground;
    //window_flags |= ImGuiWindowFlags_AlwaysAutoResize;
    //window_flags |= ImGuiWindowFlags_NoNav;
    // 开始一个新的窗口
    ImGui::Begin("Convert Window", nullptr, window_flags);

    //ImGui::SetNextItemWidth(500);
    ImGui::InputText("Xlsx Path", xlsxPath, sizeof(xlsxPath));
    ImGui::InputText("Json Path", jsonPath, sizeof(jsonPath));
    ImGui::InputText("Xlsx File", xlsxFile, sizeof(xlsxFile));
    ImGui::InputText("Json File", jsonFile, sizeof(jsonFile));
    // 在窗口中添加一个按钮
    if (ImGui::Button("xlsx2json"))
    {
        // 读取单元格的值  
        try {
            Xlsx2Json();
        }
        catch (const std::exception& e) {
            LoggerDump((std::string("Error:") += e.what()).c_str());
        }
    }
    
    ImGui::SetWindowFontScale(0.5);
    ImGui::TextWrapped("%s", logText.c_str());
    ImGui::SetWindowFontScale(1);

    // 结束窗口
    ImGui::End();
}

int main(int argc, char* argv[])
{
    std::cout << "Tool Launch Success!" << std::endl;
    // 初始化 GLFW
    if (!glfwInit())
    {
        return -1;
    }

    //ShowWindow(GetConsoleWindow(), SW_SHOWNORMAL);
    ShowWindow(GetConsoleWindow(), SW_HIDE);
    SetConsoleOutputCP(CP_UTF8);

    // 创建窗口和上下文
    GLFWwindow* window = glfwCreateWindow(800, 600, "Tool Create By ImGui", NULL, NULL);
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

    // 主循环
    while (!glfwWindowShouldClose(window))
    {
        glfwPollEvents();

        // 开始 Dear ImGui 帧
        ImGui_ImplOpenGL3_NewFrame();
        ImGui_ImplGlfw_NewFrame();
        ImGui::NewFrame();

        // 绘制简单窗口
        DrawConvertWindow();

        // 渲染 Dear ImGui
        ImGui::Render();
        int display_w, display_h;
        glfwGetFramebufferSize(window, &display_w, &display_h);
        glViewport(0, 0, display_w, display_h);
        glClearColor(0.45f, 0.55f, 0.60f, 1.00f);
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

    return 0;
}
