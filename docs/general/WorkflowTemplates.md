#### **WorkflowTemplates**



代表 **WorkflowTemplate** 对象的集合。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示当前文档中每个工作流模板的名称，然后显示特定模板的工作流特定配置用户界面。应该注意到，调用GetWorkflowTemplates方法涉及服务器的往返行程。*/ function test(){ 	const objWorkflowTemplates = Application.ActiveWorkbook.GetWorkflowTemplates(); 	for(let i = 1; i <= objWorkflowTemplates.Count; i++) 		Debug.Print(objWorkflowTemplates.Item(i).Name); 		 	var objWorkflowTemplate = objWorkflowTemplates.Item(1); 	objWorkflowTemplate.Show(); }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个代表 WorkflowTemplates 对象的容器应用程序的 Application 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 获取一个 **Long** 类型的值，指示 **WorkflowTemplates** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **WorkflowTemplates** 对象时所使用的应用程序。只读。 |

**成员属性**

#### **WorkflowTemplates.Application**

获取一个代表 WorkflowTemplates 对象的容器应用程序的 Application 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **WorkflowTemplates** 对象的变量。

#### **WorkflowTemplates.Count**

获取一个 **Long** 类型的值，指示 **WorkflowTemplates** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **WorkflowTemplates** 对象的变量。

#### **WorkflowTemplates.Creator**

获取一个 32 位整数，指示创建 **WorkflowTemplates** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **WorkflowTemplates** 对象的变量。

适用环境：web

适用平台：windows/linux