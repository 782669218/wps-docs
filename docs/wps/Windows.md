#### **Windows**



**方法**

|                                                              | 名称                         | 说明                                                         |
| ------------------------------------------------------------ | ---------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**                      | 返回一个 **Window** 对象，该对象代表文档的新窗口。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Arrange**                  | 在应用程序工作区排列所有打开的文档窗口。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BreakSideBySide**          | 如果两个窗口以并排方式排列，则终止该模式。返回一个代表该方法是否成功的 **Boolean** 类型的值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CompareSideBySideWith**    | 打开两个窗口并以并排方式排列。返回一个 **Boolean** 类型的值。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                     | 返回集合中的单个 **Window** 对象。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetPositionsSideBySide** | 重置**“并排比较”**查看模式中的两个文档窗口。                 |

**属性**

|                                                              | 名称                        | 说明                                                         |
| ------------------------------------------------------------ | --------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**             | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**                   | 返回一个 **Long** 类型的值，该值代表集合中窗口的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                 | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                  | 返回一个 **Object** 类型值，该值代表指定 **Windows** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SyncScrollingSideBySide** | 如果为 **True**，则启用同时滚动窗口内容。**Boolean** 类型，可读写。 |

**成员方法**

#### **Windows.Add**

返回一个 **Window** 对象，该对象代表文档的新窗口。

**语法**

**express.Add(Window)**

*express*   一个代表 **Windows** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Window* | 可选          | **Variant**  | 需要再打开一个窗口的 **Window** 对象。如果省略该参数，则为活动文档打开一个新窗口。 |

**说明**

当为某个文档打开多个窗口时，在窗口标题中会显示冒号 (：) 和数字。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为显示在活动窗口中的文档打开一个新窗口*/ Application.Windows.Add() /*本示例为 MyDoc.doc 打开一个新窗口*/ Application.Windows.Add(Application.Documents.Item("MyDoc.doc").Windows.Item(1))` |

#### **Windows.Arrange**

在应用程序工作区排列所有打开的文档窗口。

**语法**

**express.Arrange(ArrangeStyle)**

*express*   一个代表 **Windows** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ArrangeStyle* | 可选          | **Variant**  | 窗口排列方式。可以是下列 WdArrangeStyle 常量之一：wdIcons 或者 wdTiled。 |

**说明**

WPS 使用的是单个文档界面（SDI），本方法不会产生任何效果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例排列所有打开的窗口，使其不相互重叠。*/ Windows.Arrange(wdTiled)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将所有打开的窗口最小化，然后排列这些窗口。*/ function test() { for(let i = 1; i <= Windows.Count; i++ ){     Windows.Item(i).Activate()     Windows.Item(i).WindowState = wdWindowStateMinimize }  Windows.Arrange(wdIcons) }  ` |

#### **Windows.BreakSideBySide**

如果两个窗口以并排方式排列，则终止该模式。返回一个代表该方法是否成功的 **Boolean** 类型的值。

**语法**

**express.BreakSideBySide()**

*express*   一个代表 **Windows** 对象的变量。

**返回值**

Boolean

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例结束并排模式。*/ ActiveDocument.Windows.BreakSideBySide()` |

#### **Windows.CompareSideBySideWith**

打开两个窗口并以并排方式排列。返回一个 **Boolean** 类型的值。

**语法**

**express.CompareSideBySideWith(Document)**

*express*   一个代表 **Windows** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                   |
| ---------- | ------------- | ------------ | -------------------------- |
| *Document* | 必选          | **Variant**  | 要在并排窗口中查看的文档。 |

**返回值**

Boolean

**说明**

不能将 **Application** 对象或 **ActiveDocument** 属性与 **CompareSideBySideWith** 方法配合使用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将两个新文档排列在相邻的窗口中。*/ function test() { let objDoc1 = Documents.Add() let objDoc2 = Documents.Add()  objDoc2.Activate() objDoc2.Windows.CompareSideBySideWith (objDoc1) Windows.ResetPositionsSideBySide() }  ` |

#### **Windows.Item**

返回集合中的单个 **Window** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Windows** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 **Long** 类型值，或代表单个对象名称的 **String** 类型值。 |

#### **Windows.ResetPositionsSideBySide**

重置**“并排比较”**查看模式中的两个文档窗口。

**语法**

**express.ResetPositionsSideBySide()**

*express*   一个代表 **Windows** 对象的变量。

**说明**

对应于**“并排比较”**工具栏上的**“重设窗口位置”**按钮。使用 **ResetPositionsSideBySide** 方法可重置两个文档的显示。例如，如果用户最大化或最小化正在进行比较的两个文档窗口中的一个，则 **ResetPositionsSideBySide** 方法重置显示，从而再次并排显示两个窗口。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将以前放置在并排窗口中的两个文档放在相邻的窗口中。*/ Windows.ResetPositionsSideBySide()` |

**成员属性**

#### **Windows.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Windows** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Windows.Count**

返回一个 **Long** 类型的值，该值代表集合中窗口的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **Windows** 对象的变量。

#### **Windows.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Windows** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Windows.Parent**

返回一个 **Object** 类型值，该值代表指定 **Windows** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Windows** 对象的变量。

#### **Windows.SyncScrollingSideBySide**

如果为 **True**，则启用同时滚动窗口内容。**Boolean** 类型，可读写。

**语法**

**express.SyncScrollingSideBySide**

*express*   一个代表 **Windows** 对象的变量。

**说明**

如果该值为 **False**，则禁用同时滚动窗口。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例启用同时滚动相邻窗口的功能。*/ function test() { let objDoc1 = Documents.Add() let objDoc2 = Documents.Add()  objDoc2.Activate() objDoc2.Windows.CompareSideBySideWith(objDoc1) Windows.ResetPositionsSideBySide() Windows.SyncScrollingSideBySide = true }  ` |

适用环境：web

适用平台：windows/linux