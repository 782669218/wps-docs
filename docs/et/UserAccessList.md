**UserAccessList**



代表受保护区域用户访问权限的 **UserAccess** 对象的集合。

**说明**

使用受保护的 **Range** 对象的 **Users** 属性可返回一个 **UserAccessList** 集合。

一旦返回了 **UserAccessList** 集合，您就可以使用 **Count** 属性来确定能够访问受保护区域的用户的数量。在下例中，ET 通知用户能够访问第一个受保护区域的用户的数量。本示例假定活动工作表中存在受保护区域。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function test(){     let wksSheet = Application.ActiveSheet 									     //Notify the user the number of users that can access the protected range.     alert(wksSheet.Protection.AllowEditRanges.Item(1).Users.Count) }` |

**方法**

|                                                              | 名称          | 说明                                         |
| ------------------------------------------------------------ | ------------- | -------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**       |                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DeleteAll** | 删除有权访问工作表上受保护的区域的所有用户。 |

**属性**

|                                                              | 名称      | 说明                                           |
| ------------------------------------------------------------ | --------- | ---------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count** | 返回一个 **Long** 值，它代表集合中对象的数量。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**  | 从集合中返回一个对象。                         |

**成员方法**

#### **UserAccessList.Add**

**语法**

**express.Add(Name, AllowEdit)**

*express*   一个代表 **UserAccessList** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*      | 必选          | **String**   | 用户访问列表的名称。                                         |
| *AllowEdit* | 必选          | **Boolean**  | 如果为 True，则允许访问列表中的用户编辑受保护工作表上的可编辑单元格区域。 |

**返回值**

一个代表新的用户访问列表的 UserAccess 对象。

**说明**

添加用户访问列表。

#### **UserAccessList.DeleteAll**

删除有权访问工作表上受保护的区域的所有用户。

**语法**

**express.DeleteAll()**

*express*   一个代表 **UserAccessList** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//在本示例中，ET 将删除有权访问活动工作表中第一个受保护的区域的所有用户。本示例假定工作表不受保护。 function test() {     let wksSheet = Application.ActiveSheet 									     //Remove all users with access to the first protected range.     wksSheet.Protection.AllowEditRanges.Item(1).Users.DeleteAll() }` |

**成员属性**

#### **UserAccessList.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **UserAccessList** 对象的变量。

#### **UserAccessList.Item**

从集合中返回一个对象。

**语法**

**express.Item**

*express*   一个代表 **UserAccessList** 对象的变量。

**说明**

有关返回集合中单个成员的详细信息，请参阅返回集合中的对象。

适用环境：web

适用平台：windows/linux