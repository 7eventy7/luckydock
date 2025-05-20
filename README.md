<div align="center">

# LuckyDock
### Minimalist Rainmeter Application Docks

[![GitHub issues](https://img.shields.io/github/issues/7eventy7/luckydock.svg)](https://github.com/7eventy7/luckydock/issues)
[![License](https://img.shields.io/github/license/7eventy7/luckydock.svg)](https://github.com/7eventy7/luckydock/blob/main/LICENSE)

A minimalist and fully customizable Rainmeter skin that allows you to create sleek application docks for your desktop. It features a dedicated Python-based management menu for easy configuration and shortcut handling, letting you focus on form, function, and aesthetics.

![LuckyDockPreview](images/showcase.png)
Also featured in showcase: [ModularClocks](https://www.deviantart.com/jaxoriginals/art/ModularClocks-Clock-pack-883898019) - [YourMixer](https://forum.rainmeter.net/viewtopic.php?t=40108) - [Illustro](https://www.rainmeter.net/)

</div>

### ✨ Features
* **Customizable Docks**
  Easily tailor size, spacing, orientation, hover effects, and more through simple variable adjustments.

* **Multiple Dock Support**
  Create and manage multiple docks by duplicating and editing the provided template.

* **Smart Shortcut Handling**
  A built-in Python manager simplifies adding, editing, and organizing your application shortcuts.

* **Modern UI Design**
  Clean design with hover animations and dynamic visual feedback using ActionTimer plugin.

* **Lightweight & Responsive**
  Built on Rainmeter’s efficient skin system with optimized updates and minimal resource usage.

---

### 📝 Requirements
* Rainmeter 4.3+
* Python 3.x (for the `LuckyDockManager.pyw` script)
* ActionTimer Rainmeter plugin (for hover effects)

---

### 📦 Installation
1. Ensure you have **[Rainmeter](https://www.rainmeter.net/)** installed.
2. Extract or copy the `LuckyDock` skin folder to your Rainmeter `Skins` directory.
3. Refresh Rainmeter and load `LuckyDock1.ini` (or any other dock you create).
4. Right-click the dock → **Manage Dock** to launch the customization menu.

---

### 🧰 Managing Docks
Each dock includes a context menu item that opens the `LuckyDockManager.pyw` tool. This graphical interface allows:

* Adding new shortcuts
* Editing existing entries
* Removing apps
* Rearranging items

The dock’s appearance and behavior can be fine-tuned through the variables section in the `.ini` file (e.g., icon size, background opacity, gaps, etc.).

---

### ➕ Creating a New Dock
1. Duplicate `DockTemplate.ini` and rename it (e.g., `LuckyDock2.ini`).
2. Update the two sections at the top of the ini that feature #comments

