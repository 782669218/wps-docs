#### **WorkflowTemplate**



代表可用于当前文档的工作流之一。

**说明**

一个对应于**“启动新工作流”**对话框中显示的选项之一的 **WorkflowTemplate** 对象。在网页上，工作流模板以选项列表的形式显示。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示当前文档中每个工作流模板的名称，然后显示特定模板的工作流特定配置用户界面。应该注意到，调用 GetWorkflowTemplates 方法涉及到服务器的往返行程。*/ function test(){ 	const objWorkflowTemplates = Application.ActiveWorkbook.GetWorkflowTemplates(); 	for(let i = 1; i <= objWorkflowTemplates.Count; i++) 		Debug.Print(objWorkflowTemplates.Item(i).Name); 		 	var objWorkflowTemplate = objWorkflowTemplates.Item(1); 	objWorkflowTemplate.Show(); }` |

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Show** | 显示指定 **WorkflowTemplate** 对象的工作流特定配置用户界面。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 获取一个代表 **WorkflowTemplate** 对象的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 获取一个 32 位整数，指示创建 **WorkflowTemplate** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Description**         | 获取工作流模板的说明。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentLibraryName** | 获取与工作流模板关联的文档库的名称。只读。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentLibraryURL**  | 获取存储工作流模板的文档库的 URL 地址。只读。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**                  | 获取用于创建工作流实例的模板的 ID。只读。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 获取 **WorkflowTemplate** 对象的名称。只读。                 |

**成员方法**

#### **WorkflowTemplate.Show**

显示指定 **WorkflowTemplate** 对象的工作流特定配置用户界面。

**语法**

**express.Show()**

*express*   一个代表 **WorkflowTemplate** 对象的变量。

**返回值**

Integer

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示当前文档中每个工作流模板的名称，然后显示特定模板的工作流特定配置用户界面。*/ function test(){     const objWorkflowTemplates = Application.ActiveWorkbook.GetWorkflowTemplates();     for(let i = 1; i <= objWorkflowTemplates.Count; i++)         Debug.Print(objWorkflowTemplates.Item(i).Name);      var objWorkflowTemplate = objWorkflowTemplates.Item(1);     objWorkflowTemplate.Show(); } ` |

**成员属性**

#### **WorkflowTemplate.Application**

获取一个代表 **WorkflowTemplate** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **WorkflowTemplate** 对象的变量。

#### **WorkflowTemplate.Creator**

获取一个 32 位整数，指示创建 **WorkflowTemplate** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **WorkflowTemplate** 对象的变量。

#### **WorkflowTemplate.Description**

获取工作流模板的说明。只读。

**语法**

**express.Description**

*express*   一个代表 **WorkflowTemplate** 对象的变量。

#### **WorkflowTemplate.DocumentLibraryName**

获取与工作流模板关联的文档库的名称。只读。

**语法**

**express.DocumentLibraryName**

*express*   一个代表 **WorkflowTemplate** 对象的变量。

#### **WorkflowTemplate.DocumentLibraryURL**

获取存储工作流模板的文档库的 URL 地址。只读。

**语法**

**express.DocumentLibraryURL**

*express*   一个代表 **WorkflowTemplate** 对象的变量。

**类型**

String

#### **WorkflowTemplate.Id**

获取用于创建工作流实例的模板的 ID。只读。

**语法**

**express.Id**

*express*   一个代表 **WorkflowTemplate** 对象的变量。

#### **WorkflowTemplate.Name**

获取 **WorkflowTemplate** 对象的名称。只读。

**语法**

**express.Name**

*express*   一个代表 **WorkflowTemplate** 对象的变量。

适用环境：web

适用平台：windows/linux