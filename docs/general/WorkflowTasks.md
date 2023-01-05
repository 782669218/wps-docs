#### **WorkflowTasks**



代表 **WorkflowTask** 对象的集合。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示当前文档中每个工作流任务的名称，然后显示特定任务的工作流任务编辑用户界面。应该注意到，调用 GetWorkflowTasks 方法涉及到服务器的往返行程。*/ function test(){ 	const objWorkflowTasks = Application.ActiveWorkbook.GetWorkflowTasks(); 	for(let i = 1; i <= objWorkflowTasks.Count; i++) 		Debug.Print(objWorkflowTasks.Item(i).Name); 		 	var objWorkflowTask = objWorkflowTasks.Item(1); 	objWorkflowTask.Show(); }` |

**方法**

|                                                              | 名称     | 说明                                                        |
| ------------------------------------------------------------ | -------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 获取 **WorkflowTasks** 集合中的一个 **WorkflowTask** 对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个代表 **WorkflowTasks** 对象的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | table { word-break:break-all; }获取一个 **Long** 类型的值，指示 **WorkflowTasks** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **WorkflowTasks** 对象时所使用的应用程序。只读。 |

**成员方法**

#### **WorkflowTasks.Item**

获取 **WorkflowTasks** 集合中的一个 **WorkflowTask** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **WorkflowTasks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                             |
| -------- | ------------- | ------------ | ------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的 WorkflowTask 对象的索引号。 |

**返回值**

WorkflowTask

**成员属性**

#### **WorkflowTasks.Application**

获取一个代表 **WorkflowTasks** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **WorkflowTasks** 对象的变量。

#### **WorkflowTasks.Count**

table { word-break:break-all; }

获取一个 **Long** 类型的值，指示 **WorkflowTasks** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **WorkflowTasks** 对象的变量。

#### **WorkflowTasks.Creator**

获取一个 32 位整数，指示创建 **WorkflowTasks** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **WorkflowTasks** 对象的变量。

适用环境：web

适用平台：windows/linux