**Dialog**



代表内置的 ET 对话框。

**说明**

**Dialog** 对象是 [**Dialogs** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Dialogs/Dialogs%20.htm#jsObject_Dialogs)集合的成员。**Dialogs** 集合包含 ET 中的所有内置对话框。不能在集合中新建或添加内置对话框。**Dialog** 对象只能在 **Show** 方法中用来显示相应的对话框。

ET Visual Basic 对象库包含许多内置对话框的内置常量。每个常量的前缀均为“xlDialog”，随后即为对话框的名称。例如，**“应用名称”**对话框常量为 **xlDialogApplyNames**，**“查找文件”**对话框常量为 **xlDialogFindFile**。这些常量是 [**XlBuiltinDialog** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlBuiltInDialog%20%E6%9E%9A%E4%B8%BE.html)枚举类型的成员。

使用 **Dialogs**( *index*)（其中 *index* 是用于标识对话框的内置常量）可返回单个 **Dialog** 对象。下例运行**“文件”**菜单中的内置**“打开”**对话框。如果 ET 成功地打开了文件，则 **Show** 方法返回 **True**；如果用户取消了对话框，则返回 **False**。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let dlgAnswer = Application.Dialogs.Item(xlDialogOpen).Show()` |

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Show** | 显示内置的对话框，等待用户输入数据，然后返回一个代表用户响应的 **Boolean** 值。如果用户单击“确定”，则返回 **True**；如果用户单击“取消”，则返回 **False**。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **Dialog.Show**

显示内置的对话框，等待用户输入数据，然后返回一个代表用户响应的 **Boolean** 值。如果用户单击“确定”，则返回 **True**；如果用户单击“取消”，则返回 **False**。

**语法**

**express.Show(Arg1, ..., Arg30)**

*express*   一个代表 **Dialog** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Arg1*   | 可选          | **Variant**  | 仅应用于内置对话框，是命令的初始参数。有关详细信息，请参阅“注解”部分。 |
| *...*    | 可选          | **Variant**  | 仅应用于内置对话框，是命令的初始参数。有关详细信息，请参阅“注解”部分。 |
| *Arg30*  | 可选          | **Variant**  | 仅应用于内置对话框，是命令的初始参数。有关详细信息，请参阅“注解”部分。 |

**说明**

对于内置对话框，如果用户单击**“确定”**，则本方法返回 **True**；如果用户单击**“取消”**，则返回 **False**。

可以使用单个对话框同时更改许多属性。例如，可以使用“设置单元格格式”对话框更改 [**Font** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font)对象的所有属性。

对于一些内置对话框（如**“打开”**对话框），可以使用 *arg1*, *arg2*, ..., *arg30* 设置初始值。要查找想设置的参数，请在**内置对话框参数列表**中查找相应的对话框常量。例如，搜索 **xlDialogOpen** 常量以查找**“打开”**对话框的参数。有关内置对话框的详细信息，请参阅 [**Dialogs** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Dialogs/Dialogs%20.htm#jsObject_Dialogs)集合。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示“打开”对话框*/ Application.Dialogs.Item(xlDialogOpen).Show()` |

**成员属性**

#### **Dialog.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Dialog** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Dialog.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Dialog** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Dialog.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Dialog** 对象的变量。

适用环境：web

适用平台：windows/linux