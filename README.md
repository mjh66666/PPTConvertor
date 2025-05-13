# PPTConvertor 使用指南

![Python Version](https://img.shields.io/badge/python-3.9%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Windows Support](https://img.shields.io/badge/platform-Windows-lightgrey)

![软件图标](icons/lbd128.ico)  
*版本 1.0.0 | © 2025 萝卜丁团队*

## 🌟 项目简介


PPTConvertor 是一款专业的 PowerPoint 文件转换工具，提供以下核心功能：

- 🔄 PPT/PPTX -> 图片 -> 图片型PPT/PPTX 转换
- 🖼️ 支持 PNG/JPG 图片格式
- 🚀 批量处理能力
- 🎨 高清图标支持

---

## 💻 开发团队
- **项目作者**：莫炯豪  
- **开发成员**：萝卜丁技术团队  
- **联系方式**：mjh747786@gmail.com  
- **版权所有**：© 2025 萝卜丁工作室  

---

## 🚀 快速开始

1. **下载程序**  
   获取最新版 `PPTConvertor.exe` 可执行文件(在dist目录下)
---

## 🛠️ 开发构建
   首先确保您已经在电脑上安装了conda环境
   首先确保您已经在电脑上安装了conda
   ```bash
   # 克隆项目
   git clone https://github.com/mjh66666/PPTConvertor.git
   cd PPTConvertor
      
   # 创建Conda环境（假设环境名在yml中定义）
   conda env create -f environment.yml
   conda activate pptconvertor
   
   # 3. 安装依赖
   pip install -r requirements.txt  
   # 启动图形界面
   python src/convert_gui.py
   #打包为exe
   pyinstaller PPTConvertor.sepc
   ```
---

## 🖥️ 使用教程

#### 转换PPT三步操作
1. **选择文件**  
   - 点击"浏览"按钮选择要转换的PPT文件
   - 支持 .pptx 和 .ppt 格式

2. **设置选项**

   | 选项       | 说明                          |
   |------------|-------------------------------|
   | 图片格式   | PNG（高清）或 JPG（压缩）      |
   | 临时目录   | 建议选择空文件夹               |
   | 输出位置   | 指定生成 PPT 的保存路径        |

3. **开始转换**  
   - 点击绿色"开始转换"按钮
   - 转换过程请勿关闭PowerPoint窗口

---

## ⚠️ 重要提示
1. 转换过程中电脑不要休眠
2. 确保磁盘有足够存储空间
3. 建议关闭其他PPT文件

---

## ❓ 常见问题

### 问题1：程序打不开？
✅ 解决方案：
- 右键选择"以管理员身份运行"

### 问题2：转换失败？
✅ 排查步骤：
1. 确认PPT文件没有密码保护
2. 尝试将文件保存到桌面再操作
3. 重启程序重试

---

## 📮 技术支持
* 紧急联系  
📧 **莫炯豪**：mjh747786@gmail.com  
🏢 **萝卜丁团队**：提供专业技术支持

* 问题反馈

   创建[issue](https://github.com/mjh66666/PPTConvertor/issues)

---

*最后更新：2025年5月13号*
*© 2025 萝卜丁工作室 版权所有*