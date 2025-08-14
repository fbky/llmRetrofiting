using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api.Interop.ComApi;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Autodesk.Navisworks.Api.Automation;

namespace NavisworksHeadlessAutomation
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("Navisworks 2024+ 自动化脚本启动...");

            // 1. 解析命令行参数
            string mode = args.Length > 0 ? args[0].ToLower() : "optimize";
            string searchTerm = (args.Length > 1 && mode == "analyze") ? args[1] : "电机";

            try
            {
                // 2. 设置文件路径
                string projectDirectory = AppDomain.CurrentDomain.BaseDirectory;
                string modelFile = Path.Combine(projectDirectory, "机场三期清埗地块全地块整合.nwd");
                string outputFile = Path.Combine(projectDirectory, $"机场三期清埗地块全地块整合_Optimized_{DateTime.Now:yyyyMMddHHmmss}.fbx");

                if (!File.Exists(modelFile))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"错误：未找到模型文件：{modelFile}");
                    Console.ResetColor();
                    return;
                }
                
                // 3. 启动 Navisworks 进程 (使用新的 2024+ API)
                Console.WriteLine("\n正在后台启动 Navisworks... 这可能需要一些时间。");
                INwAutomator automator = NwAutomator.Create();
                
                // 获取底层的 COM state object
                InwOpState3 state = automator.State;

                try
                {
                    // 4. 打开文件
                    Console.WriteLine($"正在打开模型: {modelFile}");
                    state.OpenFile(modelFile);

                    // 等待文件加载完成
                    while (state.IsBusy)
                    {
                        Thread.Sleep(100);
                    }
                    Console.WriteLine("模型加载完成。");

                    // 5. 将 COM Document 转换为 .NET Document
                    Document doc = ComApiBridge.ToDocument(state.CurrentDocument);

                    // 6. 根据模式执行操作
                     switch (mode)
                    {
                        case "analyze":
                            Console.WriteLine($"\n--- 进入分析模式：搜索关键词 '{searchTerm}' ---");
                            AnalyzeModel(doc, searchTerm);
                            break;
                        
                        case "optimize":
                        default:
                            Console.WriteLine($"\n--- 进入优化模式：隐藏 '{searchTerm}' 并导出 ---");
                            OptimizeAndExportModel(doc, searchTerm, outputFile);
                            break;
                    }
                }
                finally
                {
                    // 确保进程被关闭
                    if (automator != null)
                    {
                        automator.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n--- 发生未处理的错误 ---");
                Console.WriteLine(ex.ToString());
                Console.ResetColor();
            }
            finally
            {
                Console.WriteLine("\n操作完成。按任意键退出...");
                Console.ReadKey();
            }
        }
        
        // --- 以下所有方法与之前完全相同，无需修改 ---

        private static void AnalyzeModel(Document doc, string searchTerm)
        {
            ModelItemCollection foundItems = SearchForItems(doc, searchTerm);

            if (foundItems.Count == 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"警告：在模型中未找到任何名称包含\"{searchTerm}\"的构件。");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"成功找到 {foundItems.Count} 个\"{searchTerm}\"构件。层级结构如下：");
                Console.ResetColor();
                PrintHierarchy(foundItems);
            }
        }

        private static void OptimizeAndExportModel(Document doc, string searchTerm, string outputFile)
        {
            ModelItemCollection itemsToHide = SearchForItems(doc, searchTerm);
            if (itemsToHide.Count > 0)
            {
                 Console.WriteLine($"找到 {itemsToHide.Count} 个\"{searchTerm}\"构件，将予以隐藏。");
                 doc.Models.SetHidden(itemsToHide, true);
                 Console.WriteLine("构件已在模型中隐藏。");
            }
            else
            {
                Console.WriteLine($"未找到任何\"{searchTerm}\"构件，将导出原始模型。");
            }

            Console.WriteLine("\n--- 开始导出为 FBX 文件 ---");
            PluginRecord fbxPlugin = doc.Plugins.FindExporter("lcfbx_exporter.fbx");
            if (fbxPlugin == null)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("错误：找不到FBX导出插件。");
                Console.ResetColor();
                return;
            }
            
            try 
            {
                string[] pluginOptions = doc.Plugins.GetPluginOptions(fbxPlugin.Name);
                bool optionSet = false;
                for (int i = 0; i < pluginOptions.Length; i++)
                {
                    if (pluginOptions[i].ToLower().Contains("export_hidden"))
                    {
                        pluginOptions[i+1] = "0";
                        optionSet = true;
                        break;
                    }
                }
                if(optionSet) 
                {
                     doc.Plugins.SetPluginOptions(fbxPlugin.Name, pluginOptions);
                     Console.WriteLine("已成功配置导出选项：忽略隐藏项。");
                }
                else 
                {
                    Console.WriteLine("警告：无法通过API直接设置'忽略隐藏项'。将依赖Navisworks的默认或上次手动导出设置。");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"警告：配置导出选项时发生错误 ({ex.Message})。将依赖默认设置。");
            }
            
            doc.Export(outputFile, fbxPlugin.Name);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\n--- 操作成功 ---");
            Console.WriteLine($"优化后的模型已保存至：{outputFile}");
            Console.ResetColor();
        }

        private static ModelItemCollection SearchForItems(Document doc, string searchTerm)
        {
            Search search = new Search();
            search.Selection.SelectAll();
            search.SearchConditions.Add(
                SearchCondition.HasPropertyByName("LcOaNode", "LcOaNodePropertyName")
                               .DisplayToInternal("名称")
                               .Contains(searchTerm)
            );
            return search.FindAll(doc, false);
        }

        private static void PrintHierarchy(ModelItemCollection items)
        {
            var root = new TreeNode("模型结构树");
            foreach (ModelItem item in items)
            {
                var pathParts = GetItemPathList(item);
                root.AddChild(pathParts);
            }
            root.Print();
        }

        private static List<string> GetItemPathList(ModelItem item)
        {
            var pathParts = new List<string>();
            ModelItem current = item;
            while (current != null && current.Parent != null && !current.Parent.IsRoot)
            {
                pathParts.Add(current.DisplayName);
                current = current.Parent;
            }
            pathParts.Reverse();
            return pathParts;
        }

        private class TreeNode
        {
            public string Name { get; }
            public Dictionary<string, TreeNode> Children { get; } = new Dictionary<string, TreeNode>();
            public int Count { get; set; } = 0;

            public TreeNode(string name) { Name = name; }

            public void AddChild(List<string> path)
            {
                TreeNode current = this;
                current.Count++;
                foreach (var part in path)
                {
                    if (!current.Children.ContainsKey(part))
                    {
                        current.Children[part] = new TreeNode(part);
                    }
                    current = current.Children[part];
                    current.Count++;
                }
            }

            public void Print(string indent = "", bool isLast = true)
            {
                var marker = isLast ? "└── " : "├── ";
                Console.WriteLine($"{indent}{marker}{Name} ({Count} 项)");
                indent += isLast ? "    " : "│   ";

                var lastChild = Children.Values.LastOrDefault();
                foreach (var child in Children.Values.OrderBy(c => c.Name))
                {
                    child.Print(indent, child == lastChild);
                }
            }
        }
    }
}
