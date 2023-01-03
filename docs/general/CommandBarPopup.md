#### **CommandBarPopup**



代表命令栏上的一个弹出式控件。

**说明**

每个弹出式控件都包含一个 **CommandBar** 对象。要从弹出式控件返回命令栏，请对 **CommandBar** 对象应用 **CommandBarPopup** 属性。

使用 Controls(index) 可返回一个 **CommandBarPopup** 对象，其中 *index* 是该控件的索引号。请注意，该控件的 **Type** 属性必须是 **msoControlPopup**、**msoControlGraphicPopup**、**msoControlButtonPopup**、**msoControlSplitButtonPopup** 或 **msoControlSplitButtonMRUPopup**。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

也可以使用 **FindControl** 方法返回一个 **CommandBarPopup** 对象。下面的示例在所有命令栏中搜索一个标志为“Graphics”的 **CommandBarPopup** 对象。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let myControl = Application.CommandBars.FindControl(Application.Enum.msoControlPopup, null, "Graphics")` |

**方法**

|                                                              | 名称         | 说明                                                         |
| ------------------------------------------------------------ | ------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**     | 将一个命令栏弹出式控件复制到已有的命令栏中。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**   | 将 **CommandBarPopup** 对象从其集合中删除。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Execute**  | 运行分配给指定 **CommandBarPopup** 控件的过程或内置命令。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**     | 将指定的 **CommandBarPopup** 控件移动到已有的命令栏。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset**    | 将内置 **CommandBarPopup** 控件重置为其初始功能和图符。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetFocus** | 将键盘的焦点移到指定的 **CommandBarPopup** 控件。如果弹出式控件被禁用或不可见，那么此方法将失败。 |

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**       | 获取一个 **Application** 对象，代表 **CommandBarPopup** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginGroup**        | 如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 **True**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltIn**           | 如果指定的命令栏弹出式控件是容器应用程序的内置命令栏，则为 **True**。如果是自定义命令栏，则返回 **False**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**           | 获取或设置命令栏控件的标题文字。可读写。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandBar**        | 获取一个 **CommandBar** 对象，该对象代表由指定的弹出式控件显示的菜单。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Controls**          | 获取一个代表弹出式控件上的所有控件的 **CommandBarControls** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**           | 获取一个 32 位整数，指示创建 **CommandBarPopup** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DescriptionText**   | 获取或设置命令栏弹出式控件的说明。可读写。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enabled**           | 如果启用了 **CommandBarPopup**，则此属性为 **True**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**            | 获取或设置 **CommandBarPopup** 控件的高度。可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HelpContextId**     | 获取或设置附加于 **CommandBarPopup** 控件的“帮助”主题的帮助上下文 ID 号。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HelpFile**          | 获取或设置附加于 **CommandBarPopup** 控件的“帮助”主题的文件名。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**                | 获取内置的 **CommandBarPopup** 控件的 ID。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**             | 获取一个 **Long** 类型的值，该值代表集合中 **CommandBarPopup** 对象的索引号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsPriorityDropped** | 如果当前已根据使用统计信息和布局空间将 **CommandBarPopup** 控件从菜单或工具栏中删除，则获取 **True**。（请注意，这与通过 **Visible** 属性设置的控件可见性不同）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**              | 获取指定的 **CommandBarPopup** 相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLEMenuGroup**      | 获取或设置一个 **MsoOLEMenuGroup** 常量，代表在将 OLE 服务器菜单组与 OLE 客户端菜单组合并时（即将一个容器应用程序类型的对象嵌入到另一个应用程序中时）指定的命令栏弹出式控件所属的菜单组。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLEUsage**          | 获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 **CommandBarPopup** 控件。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OnAction**          | 获取或设置一个 Visual Basic 过程的名称，当用户单击或更改**“CommandBarPopup”**控件的值时，将运行该过程。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parameter**         | 获取或设置一个应用程序可用于执行 **CommandBarPopup** 控件中命令的字符串。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**            | 获取 **CommandBarPopup** 对象的 **Parent** 对象。只读。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Priority**          | 获取或设置 **CommandBarPopup** 控件的优先级。可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tag**               | 获取或设置有关 **CommandBarPopup** 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TooltipText**       | 获取或设置 **CommandBarPopup**的**“屏幕提示”**中显示的文本。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**               | 获取指定的 **CommandBarPopup** 控件顶边到屏幕顶边的距离（以像素为单位）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**              | 获取 **CommandBarPopup** 控件的类型。只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**           | 获取或设置 **CommandBarPopup** 控件的 **Visible** 属性。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**             | 获取或设置指定的 **CommandBarPopup** 控件的宽度（以像素为单位）。可读写。 |

**成员方法**

#### **CommandBarPopup.Copy**

将一个命令栏弹出式控件复制到已有的命令栏中。

**语法**

**express.Copy(Bar, Barfore)**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Bar*     | 可选          | **Variant**  | 一个代表目标命令栏的 CommandBar 对象。如果省略此参数，控件将被复制到它原来所在的命令栏中。 |
| *Barfore* | 可选          | **Variant**  | 一个指示新控件在命令栏上的位置的数值。新控件将插入到位于此位置的控件之前。如果省略此参数，控件将被复制到命令栏的末尾。 |

#### **CommandBarPopup.Delete**

将 **CommandBarPopup** 对象从其集合中删除。

**语法**

**express.Delete(Temporary)**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Temporary* | 可选          | **Variant**  | 如果为 True，则从当前会话中删除此控件。应用程序在下次会话中将再次显示该控件。 |

**说明**

| ![img]()注释                                           |
| ------------------------------------------------------ |
| WPS Office Assistant 在 WPS Office 2015 版中已被淘汰。 |

#### **CommandBarPopup.Execute**

运行分配给指定 **CommandBarPopup** 控件的过程或内置命令。

**语法**

**express.Execute()**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本 ET 示例创建一个命令栏，然后向其中添加内置命令栏按钮控件。该按钮将执行 ET AutoSum 函数。本示例使用 Execute 方法在显示该命令栏时计算选定单元格区域的总计。 function test() {     let cbrCustBar = Application.CommandBars.Add("Custom")     let ctlAutoSum = cbrCustBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("AutoSum").Id)     cbrCustBar.Visible = true      ctlAutoSum.Execute() }` |

#### **CommandBarPopup.Move**

将指定的 **CommandBarPopup** 控件移动到已有的命令栏。

**语法**

**express.Move(Bar, Before)**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Bar*    | 可选          | **Variant**  | 一个代表控件的目标命令栏的 Command 对象。如果忽略该参数，则控件将移动到当前所在命令栏的末端。 |
| *Before* | 可选          | **Variant**  | 表示控件位置的数字。控件将插到该位置的控件之前。如果忽略该参数，控件插入到同一命令栏。 |

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将名为 Custom 的命令栏上的第一个组合框控件移动到该命令栏上的第七个控件之前。本示例将 Tag 设置为“Selection box”并赋予控件较低的优先级，以便在一行容纳不下所有控件时将其隐藏。 function test() {     let allcontrols = Application.CommandBars.Item("Custom").Controls     for(let i = 1; i <= allcontrols.Count; i++)     {         if(allcontrols.Item(i).Type == Application.Enum.msoControlComboBox)     	{             let newallcontrols = allcontrols.Item(i)             newallcontrols.Move(null,7)             newallcontrols.Tag = "Selection box"             newallcontrols.Priority = 5             break;         }     } }` |

#### **CommandBarPopup.Reset**

将内置 **CommandBarPopup** 控件重置为其初始功能和图符。

**语法**

**express.Reset()**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

重置一个内置控件将恢复该控件原来对应的动作，并将各属性恢复成初始状态。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在所有命令栏中搜索标签为“Graphics”的 CommandBarPopup 对象，然后将其重置为默认状态。 function test() {     let myControl = Application.CommandBars.FindControl(Application.Enum.msoControlPopup, null, "Graphics")      myControl.Reset() }` |

#### **CommandBarPopup.SetFocus**

将键盘的焦点移到指定的 **CommandBarPopup** 控件。如果弹出式控件被禁用或不可见，那么此方法将失败。

**语法**

**express.SetFocus()**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**成员属性**

#### **CommandBarPopup.Application**

获取一个 **Application** 对象，代表 **CommandBarPopup** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.BeginGroup**

如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 **True**。可读写。

**语法**

**express.BeginGroup**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在活动菜单栏上将最后一个控件作为一个新组的开始。 function test() {     let myMenuBar = Application.CommandBars.ActiveMenuBar     let lastMenu = myMenuBar.Controls.Item(myMenuBar.Controls.Count)     lastMenu.BeginGroup = true }` |

#### **CommandBarPopup.BuiltIn**

如果指定的命令栏弹出式控件是容器应用程序的内置命令栏，则为 **True**。如果是自定义命令栏，则返回 **False**。只读。

**语法**

**express.BuiltIn**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Caption**

获取或设置命令栏控件的标题文字。可读写。

**语法**

**express.Caption**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                             |
| ---------------------------------------- |
| 控件的标题也可作为默认的“屏幕提示”显示。 |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在自定义命令栏中添加一个带拼写检查按钮图标的命令栏控件，然后将其标题设置为“Spelling checker”。 function test() {     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)     myBar.Visible = true      let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, 2)     myControl.DescriptionText = "Starts the spelling checker"     myControl.Caption = "Spelling checker" }` |

#### **CommandBarPopup.CommandBar**

获取一个 **CommandBar** 对象，该对象代表由指定的弹出式控件显示的菜单。只读。

**语法**

**express.CommandBar**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将变量 fourthLevel 设置为名为“Drawing”的命令栏上的第四个控件。 let fourthLevel = Application.CommandBars.Item("Drawing").Controls.Item(1).CommandBar.Controls.Item(4)` |

#### **CommandBarPopup.Controls**

获取一个代表弹出式控件上的所有控件的 **CommandBarControls** 对象。只读。

**语法**

**express.Controls**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Creator**

获取一个 32 位整数，指示创建 **CommandBarPopup** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.DescriptionText**

获取或设置命令栏弹出式控件的说明。可读写。

**语法**

**express.DescriptionText**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

该说明不显示给用户，但有助于其他开发者将控件行为归档。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Enabled**

如果启用了 **CommandBarPopup**，则此属性为 **True**。可读写。

**语法**

**express.Enabled**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Height**

获取或设置 **CommandBarPopup** 控件的高度。可读写。

**语法**

**express.Height**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.HelpContextId**

获取或设置附加于 **CommandBarPopup** 控件的“帮助”主题的帮助上下文 ID 号。可读写。

**语法**

**express.HelpContextId**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

若要使用此属性，必须同时设置 HelpFile 属性。帮助主题响应 Shift+F1 按键操作。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.HelpFile**

获取或设置附加于 **CommandBarPopup** 控件的“帮助”主题的文件名。可读写。

**语法**

**express.HelpFile**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

若要使用该属性，必须同时设置 HelpContextID 属性。帮助主题响应 Shift+F1 用户按键操作。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Id**

获取内置的 **CommandBarPopup** 控件的 ID。只读。

**语法**

**express.Id**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

弹出式控件的 ID 决定了该对象的内置操作。

#### **CommandBarPopup.Index**

获取一个 **Long** 类型的值，该值代表集合中 **CommandBarPopup** 对象的索引号。只读。

**语法**

**express.Index**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

第一个命令栏控件的位置为 1。**CommandBarControls** 集合中的分隔符不计算在内。

#### **CommandBarPopup.IsPriorityDropped**

如果当前已根据使用统计信息和布局空间将 **CommandBarPopup** 控件从菜单或工具栏中删除，则获取 **True**。（请注意，这与通过 **Visible** 属性设置的控件可见性不同）。只读。

**语法**

**express.IsPriorityDropped**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

如果 **IsPriorityDropped** 为 **True**，则 **Visible** 属性设置为 **True** 的控件将不能在个性化菜单或工具栏中立即显示出来。

为了确定何时将某一特定菜单项的 **IsPriorityDropped** 属性设置为 **True**，WPS Office会记录下该菜单项的使用次数的总和，并记录下用户使用了同一菜单中的其他菜单项而没有使用该特定菜单项的应用程序会话数。当该会话数达到一定的阈值之后，便会减少使用次数总和。当总次数减为零时，便将 **IsPriorityDropped** 设置为 **True**。编程人员不能设置会话值和阈值，也不能设置 **IsPriorityDropped** 属性。但他们可以使用 **AdaptiveMenus** 属性来禁用应用程序中特定菜单的自适应菜单。

为了确定何时将某一特定工具栏控件的 **IsPriorityDropped** 属性设置为 **True**，Office 维护了一个列表，记录下了该工具栏中各个控件的最后使用顺序。工具栏会在空间允许的前提下尽可能多地显示控件，控件按使用时间由近向远排列，最近使用过的控件显示在最前面。**Priority** 设置为 1 的控件将始终显示，而且为了显示这样的控件，在必要时工具栏还将换行。编程人员可使用 **Priority** 属性来确保始终显示特定的工具栏控件，或重新定位工具栏以便使其有足够的空间显示所有的控件。

使用下表可以预测在菜单项的 **IsPriorityDropped** 属性被设置为 **True** 之前，它仍将出现在个性化菜单中的会话数。

| 命令栏控件的使用次数 | 应用程序的会话数 |
| -------------------- | ---------------- |
| 0, 1                 | 3                |
| 2                    | 6                |
| 3                    | 9                |
| 4, 5                 | 12               |
| 6- 8                 | 17               |
| 9-13                 | 23               |
| 14-24                | 29               |
| 25 及 25 以上        | 31               |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Left**

获取指定的 **CommandBarPopup** 相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。

**语法**

**express.Left**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.OLEMenuGroup**

获取或设置一个 **MsoOLEMenuGroup** 常量，代表在将 OLE 服务器菜单组与 OLE 客户端菜单组合并时（即将一个容器应用程序类型的对象嵌入到另一个应用程序中时）指定的命令栏弹出式控件所属的菜单组。可读写。

**语法**

**express.OLEMenuGroup**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

通过此属性，加载项应用程序可以指定其命令栏控件在 Office 应用程序中的表示方式。如果容器或服务器中的任何一个不实现命令栏，则执行常规的 OLE 菜单合并：菜单栏以及服务器的所有工具栏将进行合并，但容器的工具栏不会被合并。本属性只与菜单栏上的弹出式控件有关，因为菜单是根据其菜单组的类别进行合并的。

如果要合并的两个应用程序均实现了命令栏，则命令栏控件根据 **OLEUsage** 属性进行合并。

| ![img]()注释               |
| -------------------------- |
| 该属性对内置控件是只读的。 |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例检查名为“Custom”的命令栏上一个新自定义弹出式控件的 OLEMenuGroup 属性，并将该属性设置为 msoOLEMenuGroupNone。  function test() {     let myControl = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlPopup, null, null, null, false)     myControl.OLEMenuGroup = Application.Enum.msoOLEMenuGroupNone }` |

#### **CommandBarPopup.OLEUsage**

获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 **CommandBarPopup** 控件。可读写。

**语法**

**express.OLEUsage**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

通过此属性，用户可以在某 Office 应用程序与另一个 Office 应用程序合并时，指定加载项应用程序的各个命令栏控件在该 Office 应用程序中的表示方式。如果客户端和服务器均实现命令栏，则命令栏控件将逐个嵌入客户端。系统将从服务器中略去标记为只用于客户端（或既不用于客户端也不用于服务器）的自定义控件，而从客户端中略去标记为只用于服务器（或既不用于客户端也不用于服务器）的控件。剩下的控件再进行合并。

如果要合并的应用程序之一不是 Office 应用程序，则使用常规 OLE 菜单合并，该过程将由 **OLEMenuGroup** 属性控制。

#### **CommandBarPopup.OnAction**

获取或设置一个 Visual Basic 过程的名称，当用户单击或更改**“CommandBarPopup”**控件的值时，将运行该过程。可读写。

**语法**

**express.OnAction**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

容器应用程序将确定该值是否是一个合法的宏名。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Parameter**

获取或设置一个应用程序可用于执行 **CommandBarPopup** 控件中命令的字符串。可读写。

**语法**

**express.Parameter**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

如果为内置控件设置指定的参数，那么应用程序将修改其默认的操作，前提是该应用程序能分析和使用这个新值。如果为自定义控件设置参数，则可以将其用于向 Visual Basic 过程发送信息，或保存有关该控件的信息（类似于另一个 Tag 属性值）。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将一个新参数赋予控件并将焦点置于新按钮上。  function test() {     let myControl = Application.CommandBars.Item("Custom").Controls.Item(4)     myControl.Copy(null, 1)     myControl.Parameter = "2"     myControl.SetFocus() }` |

#### **CommandBarPopup.Parent**

获取 **CommandBarPopup** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Priority**

获取或设置 **CommandBarPopup** 控件的优先级。可读写。

**语法**

**express.Priority**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。无法容纳在一行中的控件将在命令栏上从右向左进行删减。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例设置命令栏弹出式控件的说明性文字和优先级。  function test() {     let popControl = Application.CommandBars.FindControl(Application.Enum.msoControlPopup, null, "Graphics")     popControl.DescriptionText = "Graphics Selection dialog"     popControl.Priority = 5 } ` |

#### **CommandBarPopup.Tag**

获取或设置有关 **CommandBarPopup** 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。

**语法**

**express.Tag**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.TooltipText**

获取或设置 **CommandBarPopup**的**“屏幕提示”**中显示的文本。可读写。

**语法**

**express.TooltipText**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

默认情况下，**Caption** 属性的值将用作**“屏幕提示”**。

#### **CommandBarPopup.Top**

获取指定的 **CommandBarPopup** 控件顶边到屏幕顶边的距离（以像素为单位）。只读。

**语法**

**express.Top**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| 注释                                                         |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Type**

获取 **CommandBarPopup** 控件的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Visible**

获取或设置 **CommandBarPopup** 控件的 **Visible** 属性。可读写。

**语法**

**express.Visible**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarPopup.Width**

获取或设置指定的 **CommandBarPopup** 控件的宽度（以像素为单位）。可读写。

**语法**

**express.Width**

*express*   一个代表 **CommandBarPopup** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

适用环境：web

适用平台：windows/linux