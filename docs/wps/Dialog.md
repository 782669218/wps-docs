**Dialog** 



代表一个内置对话框。**Dialog** 对象是 **Dialogs** 集合的成员。**Dialogs** 集合包含 WPS 中的所有内置对话框。不能在 **Dialogs** 集合中创建新的内置对话框，也不能向其中添加内置对话框。

**说明**

使用 **Dialogs**(*Index*) 返回单个 **Dialog** 对象（其中 *Index* 是标识对话框的 **WdWordDialog** 常量）。以下示例显示和执行在内置的**“打开”**对话框中采取的操作。

| 示例代码复制                                        |
| --------------------------------------------------- |
| `Application.Dialogs.Item(wdDialogFileOpen).Show()` |

**WdWordDialog** 常量由“wdDialog”前缀加上菜单和对话框的名称组成。例如，**“页面设置”**对话框的该常量为 **wdDialogFilePageSetup**，而**“新建”**对话框的该常量为 **wdDialogFileNew**。

**方法**

|                                                              | 名称        | 说明                                                         |
| ------------------------------------------------------------ | ----------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Display** | 显示指定的 WPS 内置对话框，直到用户关闭该对话框或已超过指定时限。返回一个 **Number** 类型值，指明关闭对话框时单击的按钮。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Execute** | 应用 WPS 对话框的当前设置。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Show**    | 显示和执行在指定的 WPS 内置对话框中的初始操作。返回一个 **Number** 类型值，指明关闭对话框时单击的按钮。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Update**  | 更新内置 WPS 对话框中显示的值。                              |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandBarId** | 返回一个 **Number** 类型的值，该值代表内置 WPS 对话框的工具栏控件 ID。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandName**  | 返回用来显示指定内置对话框的过程名。**String** 类型，只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTab**   | 显示指定的对话框时，该属性返回或设置活动选项卡。**WdWordDialogTab** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回一个 **Object** 类型值，该值代表指定 **Dialog** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**         | 返回内置 WPS 对话框的类型。只读 [WdWordDialog ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdWordDialog%20%E6%9E%9A%E4%B8%BE.html)类型。 |

**成员方法**

#### **Dialog.Display**

显示指定的 WPS 内置对话框，直到用户关闭该对话框或已超过指定时限。返回一个 **Number** 类型值，指明关闭对话框时单击的按钮。

**语法**

**express.Display(TimeOut)**

*express*   一个代表 **Dialog** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *TimeOut* | 可选          | **Variant**  | 自动关闭对话框前 WPS 等待的时间。每一单位大约为 0.001 秒。并发系统活动可能会增加有效时间值。如果省略这一参数，对话框只有在用户关闭它时才关闭。 |

**说明**

**Display** 方法返回下列可能值。

| 返回值    | 说明                                                     |
| --------- | -------------------------------------------------------- |
| -2        | **“关闭”**按钮。                                         |
| -1        | **“确定”**按钮。                                         |
| 0（零）   | **“取消”**按钮。                                         |
| > 0（零） | 命令按钮：1 代表第一个按钮，2 代表第二个按钮，依此类推。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示“关于”对话框 */ Application.Dialogs.Item(wdDialogHelpAbout).Display()` |

#### **Dialog.Execute**

应用 WPS 对话框的当前设置。

**语法**

**express.Execute()**

*express*   一个代表 **Dialog** 对象的变量。

#### **Dialog.Show**

显示和执行在指定的 WPS 内置对话框中的初始操作。返回一个 **Number** 类型值，指明关闭对话框时单击的按钮。

**语法**

**express.Show(TimeOut)**

*express*   一个代表 **Dialog** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *TimeOut* | 可选          | **Variant**  | 自动关闭对话框前 WPS 等待的时间。每一单位大约为 0.001 秒。并发系统活动可能会增加有效时间值。如果省略这一参数，对话框只有在用户关闭它时才关闭。 |

**说明**

下表显示 **Show** 方法所返回的值的含义。

| 返回值    | 说明                                                     |
| --------- | -------------------------------------------------------- |
| -2        | “关闭”按钮。                                             |
| -1        | “确定”按钮。                                             |
| 0（零）   | “取消”按钮。                                             |
| > 0（零） | 命令按钮：1 代表第一个按钮，2 代表第二个按钮，依此类推。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将显示并执行在“打开”对话框中启动的任意操作。文件名设置为“*.*”可显示所有文件名*/ function test() { 	let fpx = Application.Dialogs.Item(wdDialogFileOpen) 	fpx.Name = "*.*" 	fpx.Show() }` |

#### **Dialog.Update**

更新内置 WPS 对话框中显示的值。

**语法**

**express.Update()**

*express*   一个代表 **Dialog** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例返回引用“字体”对话框的 Dialog 对象，将应用于 Selection 对象的字体更改为 Arial，更新对话框的值，并显示“字体”对话框 */ function test() { 	let myDialog = Application.Dialogs.Item(wdDialogFormatFont) 	Application.Selection.Font.Name = "Arial" 	myDialog.Update() 	myDialog.Show() }` |

**成员属性**

#### **Dialog.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Dialog** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Dialog.CommandBarId**

返回一个 **Number** 类型的值，该值代表内置 WPS 对话框的工具栏控件 ID。只读。

**语法**

**express.CommandBarId**

*express*   一个代表 **Dialog** 对象的变量。

#### **Dialog.CommandName**

返回用来显示指定内置对话框的过程名。**String** 类型，只读。

**语法**

**express.CommandName**

*express*   一个代表 **Dialog** 对象的变量。

**说明**

有关使用 WPS 内置对话框的详细信息，请参阅 显示内置的 WPS 对话框。

#### **Dialog.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Dialog** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Dialog.DefaultTab**

显示指定的对话框时，该属性返回或设置活动选项卡。**WdWordDialogTab** 类型，可读写。

**语法**

**express.DefaultTab**

*express*   一个代表 **Dialog** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示所选的“页面设置”对话框中的“纸张来源”选项卡 */ function test() { 	let rng = Application.Dialogs.Item(wdDialogFilePageSetup) 	rng.DefaultTab = wdDialogFilePageSetupTabPaperSource 	rng.Show() }` |

#### **Dialog.Parent**

返回一个 **Object** 类型值，该值代表指定 **Dialog** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Dialog** 对象的变量。

#### **Dialog.Type**

返回内置 WPS 对话框的类型。只读 [WdWordDialog ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdWordDialog%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.Type**

*express*   一个代表 **Dialog** 对象的变量。

适用环境：web

适用平台：windows/linux