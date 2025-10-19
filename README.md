# ZoteroCrossRef

[![License: CC BY-NC-SA 4.0](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-blue.svg)](https://creativecommons.org/licenses/by-nc-sa/4.0/)  ![Made by Wang Nan](https://img.shields.io/badge/Made%20by-Wang%20Nan-brightgreen)

Establish cross-references between Zotero citations and the bibliography in Microsoft Word.
This VBA macro automatically creates hyperlinks from in-text citations to their corresponding bibliography entries, improving navigation and citation integrity in academic documents.

---

## 🧩 Features

* **Supported citation types:**

  * **Numeric:** `[1,3,5-9]`, `1,3,5-9`, or `[1],[3],[5-9]`, etc.
  * **Author-Year:** `(Smith, 2002; Li et al., 2025)`
* **Automatic handling of punctuation:** Compatible with both Chinese and English commas and parentheses.
* **Citation format independence:** Works regardless of whether your citation uses brackets or parentheses.
* **Non-destructive:** Does not modify the text content itself.
* **Re-runnable:** If hyperlink colors or underlines appear incorrectly, simply rerun the macro.
* **Recommended usage:**
  Run the macro **after completing the final draft**—when all citations and references are finalized—and **save a backup copy** before running it.

---

## 📂 Example Files

Example outputs are available in the `Example/` folder.
You can **click to open or download** them directly:

* [Numeric citation style example](/Example/顺序编码.pdf) [📥[Download](https://github.com/theFool-wn/ZoteroCrossRef/raw/main/Example/顺序编码.pdf)]
* [Author-year style (link year only)](/Example/作者-年（只链接年份）.pdf) [📥[Download](https://github.com/theFool-wn/ZoteroCrossRef/raw/main/Example/作者-年（只链接年份）.pdf)]
* [Author-year style (link all parts)](/Example/作者-年（全部链接）.pdf) [📥[Download](https://github.com/theFool-wn/ZoteroCrossRef/raw/main/Example/作者-年（全部链接）.pdf)]

---

## ⚙️ Usage

1. Download [`ZoteroCrossRef.bas`](https://github.com/theFool-wn/ZoteroCrossRef/raw/main/ZoteroCrossRef.bas).
2. Open and **backup** your Word document containing Zotero citations and bibliography.
3. Import and run the VBA macro `ZoteroCrossRef`.
4. Check that:

   * Each in-text citation now links to its corresponding bibliography item.
   * Formatting (color/underline) appears as expected.
5. If not, rerun the macro once more.

---

# ZoteroCrossRef（中文说明）

[![License: CC BY-NC-SA 4.0](https://img.shields.io/badge/License-CC%20BY--NC--SA%204.0-blue.svg)](https://creativecommons.org/licenses/by-nc-sa/4.0/)  ![Made by Wang Nan](https://img.shields.io/badge/Made%20by-Wang%20Nan-brightgreen)

ZoteroCrossRef 是一个用于 **在 Microsoft Word 中建立 Zotero 引文与参考文献之间交叉引用** 的 VBA 宏。
它可以自动为正文中的引文添加超链接，使其跳转到对应的参考文献条目，方便学术文档的阅读与校对。

---

## 🧩 功能特性

* **支持的引用类型：**

  * **数字型：** `[1,3,5-9]`、`1,3,5-9` 或 `[1],[3],[5-9]`，等
  * **著者-年份型：** `(Smith, 2002; Li et al., 2025)`
* **自动识别中英文标点**（括号、逗号等）。
* **不依赖特定引文样式**。
* **不改变正文内容**，仅添加跳转链接。
* **重复执行安全**，如超链接颜色或下划线不正确，可再次运行。
* **推荐使用时机：**
  建议在论文 **最终定稿后**（引文与参考文献均已确定）运行，并 **先备份文档**。

---

## 📂 示例文件

在仓库的 `Example/` 文件夹中提供了运行结果示例，点击下方链接可直接查看或下载：

* [顺序编码](./Example/顺序编码.pdf) [📥[下载](https://github.com/theFool-wn/ZoteroCrossRef/raw/main/Example/顺序编码.pdf)]
* [作者-年（只链接年份）](./Example/作者-年（只链接年份）.pdf) [📥[下载](https://github.com/theFool-wn/ZoteroCrossRef/raw/main/Example/作者-年（只链接年份）.pdf)]
* [作者-年（全部链接）](./Example/作者-年（全部链接）.pdf) [📥[下载](https://github.com/theFool-wn/ZoteroCrossRef/raw/main/Example/作者-年（全部链接）.pdf)]

---

## ⚙️ 使用方法

1. 下载 [`ZoteroCrossRef.bas`](https://github.com/theFool-wn/ZoteroCrossRef/raw/main/ZoteroCrossRef.bas)；
2. 打开并**备份**含有 Zotero 引文与参考文献的 Word 文档；
3. 载入并运行宏 `ZoteroCrossRef`；
4. 检查是否：

   * 每个引文均可跳转至对应的参考文献；
   * 超链接样式显示正确；
5. 若不正确，可重新运行一次。

---



## 🧑‍💻 Version

**Created:** Wang Nan, 2025.10.18 – 2025.10.19

**Revised:** Wang Nan, 2025.10.19

**Contact:**

* [wang.nan@buaa.edu.cn](mailto:wang.nan@buaa.edu.cn)
* [me@wangnan.net](mailto:me@wangnan.net)

**References:**

* [https://github.com/altairwei/ZoteroLinkCitation](https://github.com/altairwei/ZoteroLinkCitation)
* [https://blog.csdn.net/Bearingz/article/details/146242667](https://blog.csdn.net/Bearingz/article/details/146242667)
* [https://blog.csdn.net/eternity_memory/article/details/150343285](https://blog.csdn.net/eternity_memory/article/details/150343285)

---

## ⚖️ License

This work is licensed under the **[CC BY-NC-SA 4.0 License](https://creativecommons.org/licenses/by-nc-sa/4.0/)**.

You are free to use, share, and adapt the code for **non-commercial purposes**, provided that:

* You must give **appropriate credit**, provide **a link to this License**, and indicate if modifications were made. You may give credit in any reasonable way, but you must not do so in any way that suggests that the licensor endorses you or your use.
* You **distribute any modifications under the same license**.

© 2025 Wang Nan. All rights reserved.


