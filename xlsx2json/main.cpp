#include <iostream>
#include <random>
#include <string>
#include <iomanip>
#include <chrono>
#include <vector>
#include <fstream>
#include <sstream>
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
#include <cmath>
#include <algorithm>

#define GLFW_EXPOSE_NATIVE_WIN32
#include <GLFW/glfw3native.h>

namespace fs = std::filesystem;

// ============================ 全局变量 ============================
std::string logText = "";

// ============================ 烟花效果 ============================

/**
 * 烟花粒子类
 */
class FireworkParticle {
private:
    ImVec2 position;
    ImVec2 velocity;
    ImColor color;
    float lifeTime;
    float maxLifeTime;
    float size;
    
public:
    FireworkParticle(ImVec2 pos, ImVec2 vel, ImColor col) 
        : position(pos), velocity(vel), color(col), lifeTime(0.0f), maxLifeTime(1.5f), size(2.0f) {}
    
    void update(float deltaTime) {
        lifeTime += deltaTime;
        position.x += velocity.x * deltaTime;
        position.y += velocity.y * deltaTime;
        velocity.y += 50.0f * deltaTime; // 重力
        size *= (1.0f - deltaTime * 0.5f); // 逐渐缩小
    }
    
    void draw(ImDrawList* drawList) {
        float progress = lifeTime / maxLifeTime;
        if (progress >= 1.0f) return;
        
        float alpha = (1.0f - progress) * color.Value.w;
        float currentSize = size * (1.0f + sin(progress * 10.0f) * 0.3f); // 闪烁效果
        
        drawList->AddCircleFilled(
            position, 
            currentSize, 
            ImColor(color.Value.x, color.Value.y, color.Value.z, alpha)
        );
        
        // 拖尾效果
        ImVec2 trailPos = ImVec2(
            position.x - velocity.x * 0.1f,
            position.y - velocity.y * 0.1f
        );
        drawList->AddLine(
            position, trailPos,
            ImColor(color.Value.x, color.Value.y, color.Value.z, alpha * 0.5f),
            currentSize * 0.5f
        );
    }
    
    bool isDead() const {
        return lifeTime >= maxLifeTime;
    }
};

/**
 * 烟花类
 */
class Firework {
private:
    ImVec2 position;
    ImVec2 velocity;
    ImColor color;
    float lifeTime;
    float maxLifeTime;
    bool exploded;
    std::vector<FireworkParticle> particles;
    
public:
    Firework(ImVec2 startPos, ImVec2 targetPos, ImColor col) 
        : position(startPos), color(col), lifeTime(0.0f), maxLifeTime(3.0f), exploded(false) {
        
        // 计算初始速度指向目标位置
        ImVec2 direction = ImVec2(targetPos.x - startPos.x, targetPos.y - startPos.y);
        float distance = sqrtf(direction.x * direction.x + direction.y * direction.y);
        direction.x /= distance;
        direction.y /= distance;
        
        float speed = 150.0f + (rand() % 100); // 随机速度
        velocity = ImVec2(direction.x * speed, direction.y * speed);
    }
    
    void update(float deltaTime) {
        lifeTime += deltaTime;
        
        if (!exploded) {
            // 上升阶段
            position.x += velocity.x * deltaTime;
            position.y += velocity.y * deltaTime;
            velocity.y -= 20.0f * deltaTime; // 上升减速
            
            // 检查是否到达爆炸点或时间到
            if (velocity.y < 0 || lifeTime > maxLifeTime * 0.3f) {
                explode();
            }
        } else {
            // 爆炸后的粒子更新
            for (auto& particle : particles) {
                particle.update(deltaTime);
            }
            
            // 移除死亡的粒子
            particles.erase(
                std::remove_if(particles.begin(), particles.end(),
                    [](const FireworkParticle& p) { return p.isDead(); }),
                particles.end()
            );
        }
    }
    
    void explode() {
        if (exploded) return;
        exploded = true;
        
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_real_distribution<> angleDis(0.0, 2.0 * 3.14159);
        std::uniform_real_distribution<> speedDis(50.0, 200.0);
        std::uniform_real_distribution<> colorDis(0.7, 1.0);
        
        // 创建爆炸粒子
        int particleCount = 80 + rand() % 50;
        for (int i = 0; i < particleCount; ++i) {
            float angle = static_cast<float>(angleDis(gen));
            float speed = static_cast<float>(speedDis(gen));
            ImVec2 particleVel = ImVec2(
                cos(angle) * speed,
                sin(angle) * speed
            );
            
            // 随机颜色变化
            ImColor particleColor = ImColor(
                min(1.0f, color.Value.x * static_cast<float>(colorDis(gen))),
                min(1.0f, color.Value.y * static_cast<float>(colorDis(gen))),
                min(1.0f, color.Value.z * static_cast<float>(colorDis(gen))),
                0.8f
            );
            
            particles.emplace_back(position, particleVel, particleColor);
        }
    }
    
    void draw(ImDrawList* drawList) {
        if (!exploded) {
            // 绘制上升的烟花
            float progress = lifeTime / (maxLifeTime * 0.3f);
            float alpha = 1.0f - progress * 0.5f;
            
            drawList->AddCircleFilled(
                position, 
                3.0f, 
                ImColor(color.Value.x, color.Value.y, color.Value.z, alpha)
            );
            
            // 拖尾效果
            for (int i = 0; i < 3; ++i) {
                ImVec2 trailPos = ImVec2(
                    position.x - velocity.x * 0.02f * (i + 1),
                    position.y - velocity.y * 0.02f * (i + 1)
                );
                drawList->AddCircleFilled(
                    trailPos, 
                    2.0f - i * 0.5f, 
                    ImColor(color.Value.x, color.Value.y, color.Value.z, alpha * (0.7f - i * 0.2f))
                );
            }
        } else {
            // 绘制爆炸粒子
            for (auto& particle : particles) {
                particle.draw(drawList);
            }
        }
    }
    
    bool isDead() const {
        return exploded && particles.empty();
    }
};

/**
 * 烟花管理器
 */
class FireworkManager {
private:
    std::vector<Firework> fireworks;
    double lastFireworkTime;
    float fireworkDelay;
    
public:
    FireworkManager() : lastFireworkTime(0.0), fireworkDelay(0.2f) {}
    
    void triggerFireworks(int count, ImVec2 targetArea) {
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_real_distribution<> posDisX(0.0, 800.0);
        std::uniform_real_distribution<> posDisY(400.0, 450.0); // 从底部发射
        std::uniform_real_distribution<> colorDis(0.0, 1.0);
        
        for (int i = 0; i < count; ++i) {
            ImVec2 startPos = ImVec2(
                static_cast<float>(posDisX(gen)),
                static_cast<float>(posDisY(gen))
            );
            
            // 随机目标位置在指定区域内
            ImVec2 targetPos = ImVec2(
                targetArea.x + (rand() % 200 - 100),
                targetArea.y + (rand() % 100 - 50)
            );
            
            ImColor fireworkColor = ImColor(
                static_cast<float>(colorDis(gen)),
                static_cast<float>(colorDis(gen)),
                static_cast<float>(colorDis(gen)),
                1.0f
            );
            
            fireworks.emplace_back(startPos, targetPos, fireworkColor);
        }
    }
    
    void updateAndDraw(ImDrawList* drawList, float deltaTime) {
        // 更新所有烟花
        for (auto& firework : fireworks) {
            firework.update(deltaTime);
        }
        
        // 移除死亡的烟花
        fireworks.erase(
            std::remove_if(fireworks.begin(), fireworks.end(),
                [](const Firework& f) { return f.isDead(); }),
            fireworks.end()
        );
        
        // 绘制所有烟花
        for (auto& firework : fireworks) {
            firework.draw(drawList);
        }
    }
    
    bool isEmpty() const {
        return fireworks.empty();
    }
};

// ============================ 泡沫效果 ============================

/**
 * 碰撞特效类
 */
class CollisionEffect {
private:
    ImVec2 position;
    float lifeTime;
    float maxLifeTime;
    float radius;
    ImColor color;
    
public:
    CollisionEffect(ImVec2 pos, float size, ImColor col) 
        : position(pos), lifeTime(0.0f), maxLifeTime(0.5f), radius(size), color(col) {}
    
    void update(float deltaTime) {
        lifeTime += deltaTime;
        radius *= (1.0f + deltaTime * 2.0f); // 扩散效果
    }
    
    void draw(ImDrawList* drawList) {
        float progress = lifeTime / maxLifeTime;
        if (progress >= 1.0f) return;
        
        float alpha = (1.0f - progress) * color.Value.w;
        
        // 绘制扩散圆环
        drawList->AddCircle(
            position, 
            radius, 
            ImColor(color.Value.x, color.Value.y, color.Value.z, alpha),
            16,
            2.0f
        );
        
        // 绘制内部光晕
        drawList->AddCircleFilled(
            position, 
            radius * 0.3f, 
            ImColor(1.0f, 1.0f, 1.0f, alpha * 0.5f)
        );
    }
    
    bool isDead() const {
        return lifeTime >= maxLifeTime;
    }
};

/**
 * 泡沫类
 */
class Bubble {
private:
    ImVec2 position;
    ImVec2 velocity;
    float radius;
    ImColor color;
    float lifeTime;
    float maxLifeTime;
    float oscillation;
    float oscillationSpeed;
    bool shouldRemove;
    
public:
    Bubble() {
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_real_distribution<> posDisX(0.0, 800.0);
        std::uniform_real_distribution<> posDisY(0.0, 400.0);
        std::uniform_real_distribution<> velDis(-50.0, 50.0);
        std::uniform_real_distribution<> radiusDis(3.0, 8.0);
        std::uniform_real_distribution<> colorDis(0.7, 1.0);
        std::uniform_real_distribution<> lifeDis(15.0, 40.0);
        std::uniform_real_distribution<> oscDis(0.5, 2.0);
        
        position = ImVec2(static_cast<float>(posDisX(gen)), static_cast<float>(posDisY(gen)));
        velocity = ImVec2(static_cast<float>(velDis(gen)), static_cast<float>(velDis(gen)));
        radius = static_cast<float>(radiusDis(gen));
        
        float baseColor = static_cast<float>(colorDis(gen));
        color = ImColor(
            baseColor * 0.8f, 
            baseColor * 0.9f, 
            1.0f, 
            0.6f + static_cast<float>(colorDis(gen)) * 0.3f
        );
        
        maxLifeTime = static_cast<float>(lifeDis(gen));
        lifeTime = 0.0f;
        oscillation = 0.0f;
        oscillationSpeed = static_cast<float>(oscDis(gen));
        shouldRemove = false;
    }
    
    Bubble(float x, float y) {
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_real_distribution<> velDis(-50.0, 50.0);
        std::uniform_real_distribution<> radiusDis(3.0, 8.0);
        std::uniform_real_distribution<> colorDis(0.7, 1.0);
        std::uniform_real_distribution<> lifeDis(15.0, 40.0);
        std::uniform_real_distribution<> oscDis(0.5, 2.0);
        
        position = ImVec2(x, y);
        velocity = ImVec2(static_cast<float>(velDis(gen)), static_cast<float>(velDis(gen)));
        radius = static_cast<float>(radiusDis(gen));
        
        float baseColor = static_cast<float>(colorDis(gen));
        color = ImColor(
            baseColor * 0.8f, 
            baseColor * 0.9f, 
            1.0f, 
            0.6f + static_cast<float>(colorDis(gen)) * 0.3f
        );
        
        maxLifeTime = static_cast<float>(lifeDis(gen));
        lifeTime = 0.0f;
        oscillation = 0.0f;
        oscillationSpeed = static_cast<float>(oscDis(gen));
        shouldRemove = false;
    }
    
    void update(float deltaTime) {
        lifeTime += deltaTime;
        
        position.x += velocity.x * deltaTime;
        position.y += velocity.y * deltaTime;
        
        if (position.x - radius < 0) {
            position.x = radius;
            velocity.x = -velocity.x * 0.8f;
        } else if (position.x + radius > 800) {
            position.x = 800 - radius;
            velocity.x = -velocity.x * 0.8f;
        }
        
        if (position.y - radius < 0) {
            position.y = radius;
            velocity.y = -velocity.y * 0.8f;
        } else if (position.y + radius > 400) {
            position.y = 400 - radius;
            velocity.y = -velocity.y * 0.8f;
        }
        
        velocity.y += 20.0f * deltaTime;
        velocity.x *= (1.0f - 0.1f * deltaTime);
        velocity.y *= (1.0f - 0.1f * deltaTime);
        oscillation += oscillationSpeed * deltaTime;
    }
    
    void draw(ImDrawList* drawList) {
        float currentRadius = radius * (1.0f + 0.1f * sin(oscillation));
        float alpha = color.Value.w;
        if (lifeTime > maxLifeTime * 0.8f) {
            alpha *= 1.0f - (lifeTime - maxLifeTime * 0.8f) / (maxLifeTime * 0.2f);
        }
        
        ImColor currentColor = ImColor(color.Value.x, color.Value.y, color.Value.z, alpha);
        drawList->AddCircleFilled(position, currentRadius, currentColor);
        
        ImVec2 highlightPos = ImVec2(position.x - currentRadius * 0.3f, position.y - currentRadius * 0.3f);
        drawList->AddCircleFilled(highlightPos, currentRadius * 0.4f, ImColor(1.0f, 1.0f, 1.0f, alpha * 0.8f));
        drawList->AddCircle(position, currentRadius * 0.7f, ImColor(1.0f, 1.0f, 1.0f, alpha * 0.3f), 12, 1.5f);
        
        if (radius > 10.0f) {
            drawList->AddCircle(position, currentRadius * 0.5f, ImColor(1.0f, 1.0f, 1.0f, alpha * 0.2f), 8, 1.0f);
        }
    }
    
    bool isDead() const { return lifeTime >= maxLifeTime || shouldRemove; }
    void markForRemoval() { shouldRemove = true; }
    
    bool collideWith(Bubble& other, std::vector<CollisionEffect>& effects) {
        ImVec2 delta = ImVec2(other.position.x - position.x, other.position.y - position.y);
        float distance = sqrtf(delta.x * delta.x + delta.y * delta.y);
        float minDistance = radius + other.radius;
        
        if (distance < minDistance && distance > 0.0f) {
            Bubble* larger = (radius >= other.radius) ? this : &other;
            Bubble* smaller = (radius >= other.radius) ? &other : this;
            
            if (larger->radius > smaller->radius * 1.2f) {
                larger->radius += smaller->radius * 0.3f;
                smaller->markForRemoval();
                
                ImVec2 collisionPoint = ImVec2((position.x + other.position.x) * 0.5f, (position.y + other.position.y) * 0.5f);
                ImColor effectColor = (larger->radius > 15.0f) ? ImColor(1.0f, 0.5f, 0.2f, 0.8f) : ImColor(0.2f, 0.8f, 1.0f, 0.8f);
                effects.emplace_back(collisionPoint, larger->radius, effectColor);
                
                return true;
            } else {
                float overlap = minDistance - distance;
                ImVec2 direction = ImVec2(delta.x / distance, delta.y / distance);
                position.x -= direction.x * overlap * 0.5f;
                position.y -= direction.y * overlap * 0.5f;
                other.position.x += direction.x * overlap * 0.5f;
                other.position.y += direction.y * overlap * 0.5f;
                ImVec2 tempVel = velocity;
                velocity = other.velocity;
                other.velocity = tempVel;
                
                ImVec2 collisionPoint = ImVec2((position.x + other.position.x) * 0.5f, (position.y + other.position.y) * 0.5f);
                effects.emplace_back(collisionPoint, (radius + other.radius) * 0.5f, ImColor(1.0f, 1.0f, 1.0f, 0.6f));
            }
        }
        return false;
    }
    
    float getRadius() const { return radius; }
    ImVec2 getPosition() const { return position; }
};

/**
 * 泡沫管理器
 */
class BubbleManager {
private:
    std::vector<Bubble> bubbles;
    std::vector<CollisionEffect> effects;
    double lastTime;
    int maxBubbles;
    int lastLargeBubbleCount;
    
public:
    BubbleManager(int maxCount = 30) : maxBubbles(maxCount), lastLargeBubbleCount(0) {
        lastTime = glfwGetTime();
        for (int i = 0; i < maxCount / 2; ++i) {
            bubbles.emplace_back();
        }
    }
    
    void updateAndDraw(ImDrawList* drawList, FireworkManager& fireworkManager) {
        double currentTime = glfwGetTime();
        float deltaTime = static_cast<float>(currentTime - lastTime);
        lastTime = currentTime;
        
        for (auto& bubble : bubbles) {
            bubble.update(deltaTime);
        }
        
        std::vector<size_t> bubblesToRemove;
        for (size_t i = 0; i < bubbles.size(); ++i) {
            for (size_t j = i + 1; j < bubbles.size(); ++j) {
                if (bubbles[i].collideWith(bubbles[j], effects)) {
                    if (bubbles[i].isDead()) bubblesToRemove.push_back(i);
                    if (bubbles[j].isDead()) bubblesToRemove.push_back(j);
                }
            }
        }
        
        std::sort(bubblesToRemove.begin(), bubblesToRemove.end(), std::greater<size_t>());
        for (auto index : bubblesToRemove) {
            if (index < bubbles.size()) bubbles.erase(bubbles.begin() + index);
        }
        
        bubbles.erase(std::remove_if(bubbles.begin(), bubbles.end(), [](const Bubble& b) { return b.isDead(); }), bubbles.end());
        
        for (auto& effect : effects) {
            effect.update(deltaTime);
        }
        effects.erase(std::remove_if(effects.begin(), effects.end(), [](const CollisionEffect& e) { return e.isDead(); }), effects.end());
        
        if (bubbles.size() < maxBubbles) {
            std::random_device rd;
            std::mt19937 gen(rd());
            std::uniform_real_distribution<> addChance(0.0, 1.0);
            if (addChance(gen) < 0.1f) {
                bubbles.emplace_back();
            }
        }
        
        // 检查大泡沫数量变化并触发烟花
        int currentLargeBubbles = std::count_if(bubbles.begin(), bubbles.end(), [](const Bubble& b) { return b.getRadius() > 12.0f; });
        if (currentLargeBubbles > lastLargeBubbleCount) {
            // 新增大泡沫，触发烟花
            int newLargeBubbles = currentLargeBubbles - lastLargeBubbleCount;
            if (newLargeBubbles > 0) {
                // 找到新的大泡沫位置作为烟花目标
                for (auto& bubble : bubbles) {
                    if (bubble.getRadius() > 12.0f) {
                        fireworkManager.triggerFireworks(newLargeBubbles * 2, bubble.getPosition());
                        break; // 只在第一个大泡沫位置放烟花
                    }
                }
            }
        }
        lastLargeBubbleCount = currentLargeBubbles;
        
        // 绘制
        for (auto& effect : effects) effect.draw(drawList);
        for (auto& bubble : bubbles) bubble.draw(drawList);
        
        // 更新并绘制烟花
        fireworkManager.updateAndDraw(drawList, deltaTime);
    }
    
    void addBubble() {
        if (bubbles.size() < maxBubbles * 2) bubbles.emplace_back();
    }
    
    void addBubbleAt(float x, float y) {
        bubbles.emplace_back(x, y);
    }
    
    void setMaxBubbles(int count) { maxBubbles = count; }
};

// 全局泡沫管理器实例
std::unique_ptr<BubbleManager> bubbleManager;
std::unique_ptr<FireworkManager> fireworkManager;

// ============================ 工具函数 ============================

/**
 * 日志记录函数
 */
void LoggerDump(const char* content) {
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

/**
 * UTF-8 转 GBK
 */
std::string Utf8ToGbk(const std::string& utf8_str) {
    // UTF-8 转 UTF-16
    int wide_size = MultiByteToWideChar(CP_UTF8, 0, utf8_str.c_str(), (int)utf8_str.size(), NULL, 0);
    if (wide_size == 0) {
        return utf8_str; // 转换失败，返回原字符串
    }
    
    std::vector<wchar_t> utf16_str(wide_size + 1, 0);
    MultiByteToWideChar(CP_UTF8, 0, utf8_str.c_str(), (int)utf8_str.size(), &utf16_str[0], wide_size);
    
    // UTF-16 转 GBK
    int gbk_size = WideCharToMultiByte(CP_ACP, 0, &utf16_str[0], wide_size, NULL, 0, NULL, NULL);
    if (gbk_size == 0) {
        return utf8_str; // 转换失败，返回原字符串
    }
    
    std::vector<char> gbk_str(gbk_size + 1, 0);
    WideCharToMultiByte(CP_ACP, 0, &utf16_str[0], wide_size, &gbk_str[0], gbk_size, NULL, NULL);
    
    return std::string(&gbk_str[0]);
}

/**
 * 字符串转换：string -> wstring
 */
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

/**
 * 字符串转换：wstring -> string
 */
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

/**
 * 修改文件扩展名
 */
std::wstring changeFileExtension(const std::wstring& path, const std::wstring& newExtension) {
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

// ============================ 核心功能 ============================

template <typename Container>
std::string Join(const Container& container, const std::string& delimiter) {
    std::ostringstream oss;
    auto it = container.begin();
    if (it != container.end()) {
        oss << *it;
        ++it;
    }
    for (; it != container.end(); ++it) {
        oss << delimiter << *it;
    }
    return oss.str();
}

/**
 * Excel转JSON主函数
 */
void Xlsx2Json(std::wstring& srcPath, std::wstring& desPath) {
    LoggerDump(WcharToChar(L"=================转换开始=================").c_str());
    
    xlnt::workbook wb;
    wb.load(srcPath);
    auto ws = wb.active_sheet();
    auto dim = ws.calculate_dimension();
    
    size_t max_row = dim.height();
    size_t max_column = dim.width();
    
    LoggerDump(WcharToChar(
        std::wstring(L"列数:") + std::to_wstring(max_column) + 
        L"\t" + std::wstring(L"行数:") + std::to_wstring(max_row)
    ).c_str());
    
    Json::Value arr;
    std::vector<std::string> keys;
    
    // 处理Excel数据
    for (size_t row_index = 1; row_index <= max_row; ++row_index) {
        Json::Value tab;
        
        for (int32_t col_index = 1; col_index <= max_column; ++col_index) {
            xlnt::cell cell = ws.cell(col_index, row_index);
            
            if (row_index == 1) {
                // 第一行读取键
                keys.emplace_back(cell.to_string());
            } else {
                auto& key = keys[col_index - 1];
                if (cell.has_value())
                    tab[key] = cell.to_string();
                else
                    tab[key] = "";
            }
        }
        
        if (row_index == 1) {
            LoggerDump((std::string("[") + Join(keys, "],[") + std::string("]")).c_str());
        } else {
            arr.append(tab);
        }
    }
    
    // 写入JSON文件
    Json::StreamWriterBuilder builder;
    builder.settings_["indentation"] = "\t";
    builder.settings_["emitUTF8"] = true;
    std::unique_ptr<Json::StreamWriter> writer(builder.newStreamWriter());
    
    std::ostringstream ossJson;
    writer->write(arr, &ossJson);
    
    // 直接使用 Utf8ToGbk 转换
    std::string gbkStr = Utf8ToGbk(ossJson.str());
    
    std::ofstream outPut(desPath.c_str());
    outPut << gbkStr << std::endl;
    
    LoggerDump(WcharToChar(std::wstring(L"转换完成 ") + desPath).c_str());
    
    // 转换完成时添加一些泡沫作为庆祝效果
    if (bubbleManager) {
        for (int i = 0; i < 8; ++i) {
            bubbleManager->addBubble();
        }
    }
}

// ============================ 界面绘制 ============================

/**
 * 绘制主窗口
 */
void DrawConvertWindow() {
    ImGui::SetNextWindowPos(ImVec2(0, 0), ImGuiCond_Always);
    ImGui::SetNextWindowSize(ImVec2(800, 400));

    ImGuiWindowFlags window_flags = 0;
    window_flags |= ImGuiWindowFlags_NoMove;
    window_flags |= ImGuiWindowFlags_NoBackground;
    
    ImGui::Begin(WcharToChar(std::wstring(L"拖入xlsx文件生成json")).c_str(), nullptr, window_flags);

    // 显示日志（小字体）
    ImGui::SetWindowFontScale(0.5);
    ImGui::TextWrapped("%s", logText.c_str());
    ImGui::SetWindowFontScale(1);
    
    ImGui::End();
    
    if (bubbleManager && fireworkManager) {
        ImDrawList* drawList = ImGui::GetBackgroundDrawList();
        bubbleManager->updateAndDraw(drawList, *fireworkManager);
    }
}

// ============================ 拖拽管理 ============================

/**
 * 拖拽管理器类
 */
class DropManager : public IDropTarget {
public:
    ULONG AddRef() { return 1; }
    ULONG Release() { return 0; }

    HRESULT QueryInterface(REFIID riid, void** ppvObject) {
        if (riid == IID_IDropTarget) {
            *ppvObject = this;
            return S_OK;
        }
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }

    HRESULT DragEnter(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) {
        *pdwEffect &= DROPEFFECT_COPY;
        return S_OK;
    }

    HRESULT DragLeave() {
        return S_OK;
    }

    HRESULT DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) {
        *pdwEffect &= DROPEFFECT_COPY;
        return S_OK;
    }

    HRESULT Drop(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) {
        *pdwEffect &= DROPEFFECT_COPY;
        
        FORMATETC fmtetc = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        STGMEDIUM stgmedium;
        ZeroMemory(&stgmedium, sizeof(stgmedium));
        
        HRESULT hr = pDataObj->GetData(&fmtetc, &stgmedium);
        if (SUCCEEDED(hr)) {
            HDROP hDrop = (HDROP)stgmedium.hGlobal;
            wchar_t xlsxPath[MAX_PATH] = {0};
            DragQueryFileW(hDrop, 0, xlsxPath, MAX_PATH);
            
            try {
                std::wstring srcPath = xlsxPath;
                std::wstring desPath = changeFileExtension(xlsxPath, L".json");
                Xlsx2Json(srcPath, desPath);
            } catch (const std::exception& e) {
                LoggerDump((std::string("Error:") += e.what()).c_str());
            }
            
            DragFinish(hDrop);
            ReleaseStgMedium(&stgmedium);
        }
        
        return S_OK;
    }
};

// ============================ 主程序 ============================

int main(int argc, char* argv[]) {
    std::cout << "Tool Launch Success!" << std::endl;
    
    if (!glfwInit()) return -1;
    ShowWindow(GetConsoleWindow(), SW_HIDE);
    
    GLFWwindow* window = glfwCreateWindow(800, 400, "xlsx2json create by ImGui", NULL, NULL);
    if (!window) {
        glfwTerminate();
        return -1;
    }
    
    glfwMakeContextCurrent(window);
    glfwSwapInterval(1);
    
    if (!gladLoadGLLoader((GLADloadproc)glfwGetProcAddress)) return -1;
    
    IMGUI_CHECKVERSION();
    ImGui::CreateContext();
    ImGui_ImplGlfw_InitForOpenGL(window, true);
    ImGui_ImplOpenGL3_Init("#version 130");

    ImGuiIO& io = ImGui::GetIO();
    ImFont* font = io.Fonts->AddFontFromFileTTF("C:\\Windows\\Fonts\\msyh.ttc", 28.0f, NULL, io.Fonts->GetGlyphRangesChineseFull());
    IM_ASSERT(font != NULL);
    
    ImGuiStyle* style = &ImGui::GetStyle();
    style->Colors[ImGuiCol_Text] = ImVec4(0.0f, 1.0f, 0.0f, 1.0f);

    // 初始化管理器和烟花管理器
    bubbleManager = std::make_unique<BubbleManager>(35);
    fireworkManager = std::make_unique<FireworkManager>();

    OleInitialize(NULL);
    HWND hwnd = glfwGetWin32Window(window);
    DropManager dm;
    RegisterDragDrop(hwnd, &dm);

    while (!glfwWindowShouldClose(window)) {
        glfwPollEvents();

        ImGui_ImplOpenGL3_NewFrame();
        ImGui_ImplGlfw_NewFrame();
        ImGui::NewFrame();
        
        DrawConvertWindow();

        ImGui::Render();
        int display_w, display_h;
        glfwGetFramebufferSize(window, &display_w, &display_h);
        glViewport(0, 0, display_w, display_h);
        glClearColor(0.1f, 0.0f, 0.1f, 1.0f);
        glClear(GL_COLOR_BUFFER_BIT);
        ImGui_ImplOpenGL3_RenderDrawData(ImGui::GetDrawData());

        glfwSwapBuffers(window);
    }

    ImGui_ImplOpenGL3_Shutdown();
    ImGui_ImplGlfw_Shutdown();
    ImGui::DestroyContext();

    glfwDestroyWindow(window);
    glfwTerminate();
    
    RevokeDragDrop(hwnd);
    OleUninitialize();

    return 0;
}