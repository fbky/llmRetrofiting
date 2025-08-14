# Navisworks API 核心框架总结

本文档旨在通过总结 Navisworks .NET API 的核心功能、关键类和操作模式，为您提供一个清晰的 API 框架概览。这将帮助我们更好地规划和扩展自动化脚本，以满足模型优化的需求。

---

## 一、核心概念与三大用法

Navisworks API 提供了三种主要的与软件交互的方式：

1.  **自动化 (Automation):**
    *   **用途**: 这是我们当前项目的核心用法。它允许从外部应用程序（如我们正在编写的控制台程序）来驱动 Navisworks。
    *   **优势**: 可以在后台执行任务，无需打开 Navisworks 用户界面，非常适合批处理和自动化流程，例如打开文件、执行搜索、隐藏构件和导出。
    *   **关键类**: `NavisworksApplication`

2.  **插件 (Plug-Ins):**
    *   **用途**: 创建直接集成在 Navisworks 软件内部的新功能或工具（例如，在顶部菜单栏添加一个自定义按钮）。
    *   **优势**: 可以与用户界面深度集成，创建交互式工具。
    *   **关键接口**: `AddInPlugin`, `DockablePlugin`

3.  **控件 (Controls):**
    *   **用途**: 将 Navisworks 的模型查看器功能嵌入到您自己的独立应用程序中。
    *   **优势**: 可以构建包含 3D 视图的自定义软件。

---

## 二、关键 API 对象与操作

以下是在自动化场景中，我们最需要关注的 API 类和它们的作用：

| 核心类/对象                 | 功能描述                                                                                                       | 关键属性/方法                                                                                                                                                             |
| --------------------------- | -------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `NavisworksApplication`     | **应用程序本身**。代表一个 Navisworks 实例，是所有自动化操作的入口点。                                           | `new NavisworksApplication(path)` (构造函数), `.OpenFile(filePath)`, `.ActiveDocument`                                                                                         |
| `Document`                  | **模型文档**。代表一个打开的 `.nwd` 或 `.nwf` 文件，是进行模型操作的总控制器。                                   | `.Models` (获取所有追加的模型), `.CurrentFileName` (获取文件名), `.Plugins` (访问插件管理器), `.Export(...)` (执行导出)                                                            |
| `ModelItem`                 | **模型构件**。代表模型中的一个具体元素，如一根管道、一个阀门或一个楼层。它是我们进行查找、读取和操作的最小单位。 | `.DisplayName` (获取构件名称), `.Parent` (获取父节点), `.Children` (获取子节点), `.PropertyCategories` (访问所有属性)                                                        |
| `Search`                    | **搜索工具**。用于在模型中根据特定条件查找一组 `ModelItem` 构件。                                                | `new Search()` (创建搜索), `.SearchConditions.Add(...)` (添加搜索条件), `.FindAll(doc, ...)` (执行搜索)                                                                           |
| `SearchCondition`           | **搜索条件**。定义了查找构件时依据的规则，例如根据名称、类型、GUID 或任何其他属性进行筛选。                          | `SearchCondition.HasPropertyByName(...)`, `.DisplayToInternal(...)` (将显示名称转换为内部名称), `.Contains(...)` (包含条件), `.Equal(...)` (等于条件) |
| `PropertyCategory`          | **属性类别**。代表一个属性分组，例如“项目”、“尺寸”、“材质”等。每个 `ModelItem` 都包含多个属性类别。           | `.DisplayName` (类别名称), `.Properties` (该类别下的所有属性)                                                                                                                 |
| `DataProperty`              | **具体属性**。代表一个键值对，例如 `名称: "风管"` 或 `长度: 1500mm`。                                            | `.DisplayName` (属性名), `.Value` (属性值, 类型为 `VariantData`)                                                                                                          |
| `PluginRecord` & `Plugins`  | **插件管理器**。用于查找和配置 Navisworks 的内置功能，例如 FBX 导出器。                                          | `doc.Plugins.FindExporter(...)` (查找导出器), `.SetPluginOptions(...)` (设置插件选项，如“忽略隐藏项”)                                                                    |

---

## 三、核心操作流程（我们的用例）

结合以上 API，我们可以总结出探索和优化模型的典型步骤：

1.  **初始化与加载**:
    *   创建一个 `NavisworksApplication` 实例。
    *   使用 `.OpenFile()` 方法加载目标 `.nwd` 文件，获取 `Document` 对象。

2.  **探索与分析 (新功能)**:
    *   **遍历模型树**: 从 `doc.Models.RootItems` 开始，递归访问每个 `ModelItem` 的 `.Children` 属性，可以完整地构建出整个模型的层级结构。
    *   **分类汇总**: 遍历过程中，读取每个 `ModelItem` 的 `DisplayName` 和属性，按类型（如风管、桥架、设备）进行分类和计数。这能帮我们识别出哪些构件占比较高，是优化的重点。
    *   **属性查询**: 对特定的 `ModelItem`，访问其 `PropertyCategories` 和 `DataProperty`，可以读取到详细的参数信息（如尺寸、材质、来源文件等）。

3.  **筛选与操作 (现有功能)**:
    *   **创建 `Search` 对象**，并定义一个或多个 `SearchCondition` 来精确查找需要处理的构件（例如，所有名称包含“支架”且材质为“钢”的构件）。
    *   **执行搜索**，用 `FindAll` 得到一个 `ModelItemCollection`。
    *   **执行修改**: 对这个集合执行操作，例如调用 `doc.Models.SetHidden(collection, true)` 将它们隐藏。

4.  **导出**:
    *   使用 `doc.Plugins.FindExporter("lcfbx_exporter.fbx")` 找到 FBX 导出器。
    *   **（关键）** 配置导出选项，确保“忽略隐藏项”被设置。
    *   调用 `doc.Export()` 方法，将优化后的模型保存为新的 `.fbx` 文件。

---

现在，我们对 Navisworks API 的能力有了全面的了解。下一步，我将基于这个框架，修改 `RunAutomation.cs` 文件，为它增加一个强大的**“分析模式”**，让我们可以通过输入不同的关键词来深入探索模型结构。

您准备好进行下一步了吗？
