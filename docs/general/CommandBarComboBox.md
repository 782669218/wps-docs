#### **CommandBarComboBox**



代表命令栏中的一个组合框控件。

**说明**

使用 **Controls(index)** 可返回一个 **CommandBarComboBox** 对象，其中 *index* 为该控件的索引号。请注意，该控件的 **Type** 属性必须是 **msoControlEdit**、**msoControlDropdown**、**msoControlComboBox**、**msoControlButtonDropdown**、**msoControlSplitDropdown**、**msoControlOCXDropdown**、**msoControlGraphicCombo** 或 **msoControlGraphicDropdown**。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**方法**

|                                                              | 名称           | 说明                                                         |
| ------------------------------------------------------------ | -------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddItem**    | 向指定命令栏组合框控件添加列表项。该组合框控件必须是自定义控件，并且必须是下拉列表框或组合框。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Clear**      | 从命令栏组合框控件（下拉列表框或组合框）中删除所有列表项。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**       | 将一个命令栏组合框控件复制到已有的命令栏中。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**     | 将 **CommandBarCombo** 控件对象从其集合中删除。在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Execute**    | 运行分配给指定 **CommandBarComboBox** 控件的过程或内置命令。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**       | 将指定的控件移动到已有的命令栏。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RemoveItem** | 从 **CommandBarComboBox** 控件中删除一项。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset**      | 将内置命令栏重置为其默认配置，或将内置 **CommandBarComboBox** 控件重置为其初始功能和图符。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetFocus**   | 将键盘的焦点移到指定 **CommandBarComboBox** 控件。如果该控件被禁用或不可见，那么此方法将失败。 |

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**       | 获取一个 **Application** 对象，代表 **CommandBarComboBox** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BeginGroup**        | 如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 True。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltIn**           | 如果指定的命令栏控件是容器应用程序的内置控件，则获取 **True**。如果是自定义控件或是已设置了 **OnAction** 属性的内置控件，则返回 **False**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**           | 获取或设置命令栏控件的标题文字。可读写。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**           | 获取一个 32 位整数，指示创建 **CommandBarComboBox** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DescriptionText**   | 获取或设置命令栏组合框控件的说明。可读写。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DropDownLines**     | 获取或设置命令栏组合框控件中的行数。该组合框控件必须为自定义控件并且是下拉列表框或组合框。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DropDownWidth**     | 获取或设置指定命令栏组合框控件的列表宽度（以像素为单位）。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enabled**           | 获取或设置用于指定是否启用 **CommandBarComboBox** 的 **Boolean** 值。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**            | 获取或设置 **CommandBarComboBox** 控件的高度。可读写。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HelpContextId**     | 获取或设置附加于 **CommandBarComboBox** 控件的“帮助”主题的帮助上下文 ID 号。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HelpFile**          | 获取或设置附加于 **CommandBarComboBox** 控件的“帮助”主题的文件名。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**                | 获取内置的 **CommandBarComboBox** 控件的 ID。只读。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**             | 获取一个 **Long** 类型的值，该值代表集合中 **CommandBarComboBox** 对象的索引号。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsPriorityDropped** | 如果当前已根据使用统计信息和布局空间将控件从菜单或工具栏中删除，则返回 **True**。（请注意，这与通过 **Visible** 属性设置的控件可见性不同）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**              | 获取 **CommandBarComboBox** 控件相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **List**              | 获取或设置 **CommandBarComboBox** 控件中的某一项。可读/写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListCount**         | 获取 **CommandBarComboBox** 控件中列表项的数目。只读。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListHeaderCount**   | 获取或设置出现在分隔行之上的 **CommandBarComboBox** 控件中的列表项数目。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListIndex**         | 获取或设置 **CommandBarComboBox** 控件的列表部分中选定项的索引号。如果没有选择列表中的任何项，此属性将返回 0。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLEUsage**          | 获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 **CommandBarComboBox** 控件。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OnAction**          | 获取或设置一个 Visual Basic 过程的名称，该过程在用户单击或更改 **CommandBarComboBox** 控件的值时运行。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parameter**         | 获取或设置一个应用程序可用于执行 **CommandBarComboBox** 控件中命令的字符串。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**            | 获取 **CommandBarComboBox** 对象的 **Parent** 对象。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Priority**          | 获取或设置 **CommandBarComboBox** 控件的优先级。对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**             | 获取或设置 **CommandBarComboBox** 控件的显示方式。该属性值可设置为以下 **MsoComboStyle** 常量之一：**msoComboLabel** 或 **msoComboNormal**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tag**               | 获取或设置有关 **CommandBarComboBox** 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**              | 获取或设置 **CommandBarComboBox** 控件的显示或编辑区中的文字。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TooltipText**       | 获取或设置 **CommandBarComboBox**的**“屏幕提示”**中显示的文本。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**               | 获取指定的 **CommandBarComboBox** 控件顶边到屏幕顶边的距离（以像素为单位）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**              | 获取 **CommandBarComboBox** 控件的类型。只读。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**           | 获取或设置 **CommandBarComboBox** 控件的 **Visible** 属性。如果 **CommandBarControl** 可见，则此属性为 **True**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**             | 获取或设置指定的 **CommandBarComboBox** 控件的宽度（以像素为单位）。可读写。 |

**成员方法**

#### **CommandBarComboBox.AddItem**

向指定命令栏组合框控件添加列表项。该组合框控件必须是自定义控件，并且必须是下拉列表框或组合框。

**语法**

**express.AddItem(Text, Index)**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                   |
| -------- | ------------- | ------------ | ---------------------------------------------------------- |
| *Text*   | 必选          | **String**   | 要添加到控件的文字。                                       |
| *Index*  | 可选          | **Variant**  | 项在列表中的位置。如果省略该参数，项会被添加到列表的尾部。 |

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| ![img]()注释                             |
| ---------------------------------------- |
| 此方法不能用于编辑框或内置的组合框控件。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在命令栏中添加一个组合框控件，然后在该控件中添加两项。本示例还设置项的行数和组合框的宽度。  function test() {     let myBar = Application.CommandBars.Item("Custom")     let myControl = myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)     myControl.AddItem("First Item", 1)     myControl.AddItem("Second Item", 2)     myControl.DropDownLines = 3     myControl.DropDownWidth = 75     myControl.ListHeaderCount = 0 }` |

#### **CommandBarComboBox.Clear**

从命令栏组合框控件（下拉列表框或组合框）中删除所有列表项。

**语法**

**express.Clear()**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 如果将此方法应用于内置命令栏控件（WPS Office附带的控件），此方法将失败。 |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将检查名为“自定义栏”的命令栏上组合框控件中的项目数。如果组合框的列表中的项目数少于 3，那么本示例清除该列表，在其中添加新的第一项，然后将该新项目显示为组合框控件的默认值。 function test() {     let myBar = Application.CommandBars.Item("Custom Bar")     let myControl = myBar.Controls.Item(Application.Enum.msoControlComboBox)     myControl.AddItem("First Item", 1)     myControl.AddItem("Second Item", 2)     if(myControl.ListCount < 3)     {         myControl.Clear()         myControl.AddItem("New Item", 1)     } }` |

#### **CommandBarComboBox.Copy**

将一个命令栏组合框控件复制到已有的命令栏中。

**语法**

**express.Copy(Bar, Before)**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Bar*    | 可选          | **Variant**  | 一个代表目标命令栏的 CommandBar 对象。如果省略此参数，控件将被复制到它原来所在的命令栏中。 |
| *Before* | 可选          | **Variant**  | 一个指示新控件在命令栏上位置的数值。新控件将插入到位于此位置的控件之前。如果省略此参数，控件将被复制到命令栏的末尾。 |

**返回值**

CommandBarControl

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Delete**

将 **CommandBarCombo** 控件对象从其集合中删除。在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。

**语法**

**express.Delete(Temporary)**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Temporary* | 可选          | **Variant**  | 如果为 True，则从当前会话中删除此控件。应用程序在下次会话中将再次显示该控件。 |

#### **CommandBarComboBox.Execute**

运行分配给指定 **CommandBarComboBox** 控件的过程或内置命令。

**语法**

**express.Execute()**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本 ET 示例创建一个命令栏，然后向其中添加内置命令栏按钮控件。该按钮将执行 ET AutoSum 函数。本示例使用 Execute 方法在显示该命令栏时计算选定单元格区域的总计。 function test() {     let cbrCustBar = Application.CommandBars.Add("Custom")     let ctlAutoSum = cbrCustBar.Controls.Add(Application.Enum.msoControlButton, CommandBars.Item("Standard").Controls.Item("AutoSum").Id)     cbrCustBar.Visible = true      ctlAutoSum.Execute() }` |

#### **CommandBarComboBox.Move**

将指定的控件移动到已有的命令栏。

**语法**

**express.Move(Bar, Before)**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

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
| `//本示例将名为 Custom 的命令栏上的第一个组合框控件移动到该命令栏上的第七个控件之前。本示例将 Tag 设置为“Selection box”并赋予控件较低的优先级，以便在一行容纳不下所有控件时将其隐藏。 function test() {     let allcontrols = Application.CommandBars.Item("Custom").Controls     for(let i = 1; i <= allControls.Count; i++)     {         if(allControls.Item(i).Type == Application.Enum.msoControlComboBox)     	{             allControls.Item(i).Move(undefined,7)             allControls.Item(i).Tag = "Selection box"             allControls.Item(i).Priority = 5         }     } }` |

#### **CommandBarComboBox.RemoveItem**

从 **CommandBarComboBox** 控件中删除一项。

**语法**

**express.RemoveItem()**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| ![img]()注释               |
| -------------------------- |
| 本属性只能用于列表框控件。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例确定指定组合框所含的列表项是否多于三项。如果多于三项，那么本示例将删除第二项，更改样式并设置一个新值。此外，本示例还设置了父对象（即 CommandBarControl 对象）的 Tag 属性以表明该列表发生了变化。 function test() {     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop,undefined, true)     myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)     myBar.Visible = true      let myControls = Application.CommandBars.Item("Custom").Controls.Item(1)     myControls.AddItem("Get Stock Quote", 1)     myControls.AddItem("View Chart", 2)     myControls.AddItem("View Fundamentals", 3)     myControls.AddItem("View News", 4)     myControls.Caption = "Stock Data"     myControls.DescriptionText = "View Data For Stock"     let myControl = myBar.Controls.Item(1)     if(myControl.ListCount > 3)     {        myControl.RemoveItem(2)        myControl.Style = Application.Enum.msoComboNormal        myControl.Text = "New Default"        let ctrl = myControl.Parent     } }` |

#### **CommandBarComboBox.Reset**

将内置命令栏重置为其默认配置，或将内置 **CommandBarComboBox** 控件重置为其初始功能和图符。

**语法**

**express.Reset()**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

重置一个内置控件将恢复该控件原来对应的动作，并将各属性恢复成初始状态。重置一个内置命令栏将删除其中的自定义控件并恢复其内置控件。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例自定义一个命令栏组合框。首先，将组合框重置为其默认状态。然后将两个行项添加到组合框，并设置各种属性。 function test() {     let combo = Application.CommandBars.Item("Custom").Controls.Item(2)     combo.Reset()     combo.AddItem("First Item", 1)     combo.AddItem("Second Item", 2)     combo.DropDownLines = 3     combo.DropDownWidth = 75     combo.ListIndex = 0  }` |

#### **CommandBarComboBox.SetFocus**

将键盘的焦点移到指定 **CommandBarComboBox** 控件。如果该控件被禁用或不可见，那么此方法将失败。

**语法**

**express.SetFocus()**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例创建一个名为“Custom”的命令栏，并在其中添加一个 ComboBox 控件和一个 Button 控件，然后使用 SetFocus 方法为 ComboBox 控件设置焦点。 function test() {     let focusBar = Application.CommandBars.Add("Custom")     CommandBars.Item("Custom").Visible = true      CommandBars.Item("Custom").Position = Application.Enum.msoBarTop     									     let testComboBox = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlComboBox, 1)     testComboBox.AddItem("First Item", 1)     testComboBox.AddItem("Second Item", 2)     let testButton = Application.CommandBars.Item("Custom").Controls.Add(Application.Enum.msoControlButton)     testButton.FaceId = 17     // Set the focus to the combo box.     testComboBox.SetFocus() }` |

**成员属性**

#### **CommandBarComboBox.Application**

获取一个 **Application** 对象，代表 **CommandBarComboBox** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.BeginGroup**

如果指定的命令栏控件显示在命令栏上控件组的最前面，则获取 True。可读写。

**语法**

**express.BeginGroup**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.BuiltIn**

如果指定的命令栏控件是容器应用程序的内置控件，则获取 **True**。如果是自定义控件或是已设置了 **OnAction** 属性的内置控件，则返回 **False**。只读。

**语法**

**express.BuiltIn**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Caption**

获取或设置命令栏控件的标题文字。可读写。

**语法**

**express.Caption**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                             |
| ---------------------------------------- |
| 控件的标题也可作为默认的“屏幕提示”显示。 |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Creator**

获取一个 32 位整数，指示创建 **CommandBarComboBox** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.DescriptionText**

获取或设置命令栏组合框控件的说明。可读写。

**语法**

**express.DescriptionText**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

该说明不显示给用户，但有助于其他开发者将控件行为归档。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.DropDownLines**

获取或设置命令栏组合框控件中的行数。该组合框控件必须为自定义控件并且是下拉列表框或组合框。可读写。

**语法**

**express.DropDownLines**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

如果将此属性设置为 0（零），则控件中的行数将取决于列表中的项数。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 如果试图设置该属性的组合框控件是编辑框或内置组合框控件，则会出现错误。 |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在命令栏“Custom”中添加一个包含两项的组合框控件。此外，本示例还设置项的行数，组合框的宽度，并将该组合框的默认值设置为空。  function test() {     let myBar = Application.CommandBars.Item("Custom")     let myControl = myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)									     myControl.AddItem ("First Item", 1)     myControl.AddItem ("Second Item", 2)     myControl.DropDownLines = 3     myControl.DropDownWidth = 75     myControl.ListHeaderCount = 0 }` |

#### **CommandBarComboBox.DropDownWidth**

获取或设置指定命令栏组合框控件的列表宽度（以像素为单位）。可读写。

**语法**

**express.DropDownWidth**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

如果将该属性设置为 -1，那么将根据组合框列表中最长的项来设置列表的宽度。如果将该属性设置为 0，那么将根据控件的宽度来设置列表的宽度。

| ![img]()注释                                     |
| ------------------------------------------------ |
| 如果为内置的组合框控件设置该属性，则会出现错误。 |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在命令栏“Custom”中添加一个包含两项的组合框控件。此外，本示例还设置项的行数，组合框的宽度，并将该组合框的默认值设置为空。  function test() {     let myBar = Application.CommandBars.Item("Custom")     let myControl = myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)									     myControl.AddItem ("First Item", 1)     myControl.AddItem ("Second Item", 2)     myControl.DropDownLines = 3     myControl.DropDownWidth = 75     myControl.ListHeaderCount = 0 }` |

#### **CommandBarComboBox.Enabled**

获取或设置用于指定是否启用 **CommandBarComboBox** 的 **Boolean** 值。可读写。

**语法**

**express.Enabled**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

对于命令栏，如果将该属性设置为 **True**，则该命令栏的名称将出现在可用命令栏列表中。

对于内置控件，如果将 **Enabled** 属性设置为 **True**，则将由应用程序确定其状态。但如果将该属性设置为 **False**，则强行禁用该控件。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Height**

获取或设置 **CommandBarComboBox** 控件的高度。可读写。

**语法**

**express.Height**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在名为 Custom 的命令栏上添加一个自定义控件。本示例设置该自定义控件的高度为命令栏高度的两倍，并设置该控件宽度为 50 像素。请注意命令栏如何根据控件的大小自我调整。  function test() {     let myBar = Application.CommandBars.Item("Custom")     let barHeight = myBar.Height     let myControl = myBar.Controls.Add(Application.Enum.msoControlButton,CommandBars.Item("Standard").Controls.Item("Save").Id,null,null,true)     myControl.Height = barHeight * 2     myControl.Width = 50     myBar.Visible = true }` |

#### **CommandBarComboBox.HelpContextId**

获取或设置附加于 **CommandBarComboBox** 控件的“帮助”主题的帮助上下文 ID 号。可读写。

**语法**

**express.HelpContextId**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

若要使用此属性，必须同时设置 HelpFile 属性。帮助主题响应 Shift+F1 按键操作。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例添加了一个自定义命令栏以及跟踪股票的组合框。该示例也指定了在用户按下 Shift+F1 时所显示的组合框“帮助”主题。  function test() {     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)     myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)     myBar.Visible = true                             let cont = CommandBars.Item("Custom").Controls.Item(1)     cont.AddItem("Get Stock Quote", 1)     cont.AddItem("View Chart", 2)     cont.AddItem("View Fundamentals", 3)     cont.AddItem("View News", 4)     cont.Caption = "Stock Data"     cont.DescriptionText = "View Data For Stock"     cont.HelpFile = "C:\\corphelp\\custom.hlp"     cont.HelpContextID = 47 }` |

#### **CommandBarComboBox.HelpFile**

获取或设置附加于 **CommandBarComboBox** 控件的“帮助”主题的文件名。可读写。

**语法**

**express.HelpFile**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

若要使用该属性，必须同时设置 HelpContextID 属性。帮助主题响应 Shift+F1 用户按键操作。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例添加了一个自定义命令栏以及跟踪股票的组合框。该示例也指定了在用户按下 Shift+F1 时所显示的组合框“帮助”主题。  function test() {     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)     myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)     myBar.Visible = true                             let cont = CommandBars.Item("Custom").Controls.Item(1)     cont.AddItem("Get Stock Quote", 1)     cont.AddItem("View Chart", 2)     cont.AddItem("View Fundamentals", 3)     cont.AddItem("View News", 4)     cont.Caption = "Stock Data"     cont.DescriptionText = "View Data For Stock"     cont.HelpFile = "C:\\corphelp\\custom.hlp"     cont.HelpContextID = 47 }` |

#### **CommandBarComboBox.Id**

获取内置的 **CommandBarComboBox** 控件的 ID。只读。

**语法**

**express.Id**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

控件的 ID 号决定了该控件的内置操作。所有自定义控件的 **Id** 属性的值都是 1。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| WPS Office Assistant 在 WPS Office system 2015 版中已被淘汰。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例将名为“Standard”的工具栏上每个控件的标题更改为该控件当前的 Id 属性值。  function test() {     let ctl = Application.CommandBars.Item("Standard").Controls     for(let i=1; i<=ctl.Count; i++)     {         ctl.Item(i).Caption = String(ctl.Id)     } }` |

#### **CommandBarComboBox.Index**

获取一个 **Long** 类型的值，该值代表集合中 **CommandBarComboBox** 对象的索引号。只读。

**语法**

**express.Index**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

第一个命令栏控件的位置为 1。**CommandBarControls** 集合中的分隔符不计算在内。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在名为“Custom”的命令栏中搜索 Id 值为 23 的控件。如果能找到这样的控件，并且其索引号大于 5，则将该控件放置在命令栏上的第一个位置上。 function test() {     let myBar = CommandBars.Item("Custom")     let ctrl1 = myBar.FindControl(null,23)     if (ctrl1 && ctrl1.Index > 5)     {         ctrl1.Move (null,1)     } }` |

#### **CommandBarComboBox.IsPriorityDropped**

如果当前已根据使用统计信息和布局空间将控件从菜单或工具栏中删除，则返回 **True**。（请注意，这与通过 **Visible** 属性设置的控件可见性不同）。只读。

**语法**

**express.IsPriorityDropped**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

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

#### **CommandBarComboBox.Left**

获取 **CommandBarComboBox** 控件相对于屏幕左边缘的水平位置（以像素为单位）。返回距离停靠区域左侧的距离。只读。

**语法**

**express.Left**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.List**

获取或设置 **CommandBarComboBox** 控件中的某一项。可读/写。

**语法**

**express.List**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

| ![img]()注释                         |
| ------------------------------------ |
| 对于内置的组合框控件，该属性为只读。 |

**示例**

#### **CommandBarComboBox.ListCount**

获取 **CommandBarComboBox** 控件中列表项的数目。只读。

**语法**

**express.ListCount**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

#### **CommandBarComboBox.ListHeaderCount**

获取或设置出现在分隔行之上的 **CommandBarComboBox** 控件中的列表项数目。可读写。

**语法**

**express.ListHeaderCount**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

**ListHeaderCount** 属性值为 -1 表示组合框控件中没有分隔行。

| ![img]()注释                         |
| ------------------------------------ |
| 对于内置的组合框控件，该属性为只读。 |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

#### **CommandBarComboBox.ListIndex**

获取或设置 **CommandBarComboBox** 控件的列表部分中选定项的索引号。如果没有选择列表中的任何项，此属性将返回 0。可读写。

**语法**

**express.ListIndex**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

设置 **ListIndex** 属性会使指定控件选择给定的项并在应用程序中执行与相应的操作。

| ![img]()注释             |
| ------------------------ |
| 此属性只能用于列表控件。 |

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

#### **CommandBarComboBox.OLEUsage**

获取或设置特定的 OLE 客户端和 OLE 服务器角色，在合并两个 WPS Office应用程序时，会用到这些角色中的 **CommandBarComboBox** 控件。可读写。

**语法**

**express.OLEUsage**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

本属性允许用户在某 Office 应用程序与另一个 Office 应用程序合并时，指定加载应用程序的各个命令栏控件在该 Office 应用程序中的表示方法。如果客户端和服务器均提供命令栏，那么命令栏控件将逐个嵌入客户端。系统将从服务器中略去标记为只用于客户端（或既不用于客户端也不用于服务器）的自定义控件，而从客户端中略去标记为只用于服务器（或既不用于客户端也不用于服务器）的控件。剩下的控件再进行合并。

如果要合并的应用程序之一不是 Office 应用程序，则使用常规 OLE 菜单合并，该过程将由 OLEMenuGroup 属性控制。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.OnAction**

获取或设置一个 Visual Basic 过程的名称，该过程在用户单击或更改 **CommandBarComboBox** 控件的值时运行。可读写。

**语法**

**express.OnAction**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

容器应用程序将确定该值是否是一个合法的宏名。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Parameter**

获取或设置一个应用程序可用于执行 **CommandBarComboBox** 控件中命令的字符串。可读写。

**语法**

**express.Parameter**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

如果为内置控件设置一个指定的参数，那么应用程序将修改其默认的操作，前提是该应用程序能分析和使用这个新值。如果为自定义控件设置参数，那么该参数可用来为 Visual Basic 程序传送信息，或保存关于该控件的信息（类似于另一个 Tag 属性值）。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Parent**

获取 **CommandBarComboBox** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Priority**

获取或设置 **CommandBarComboBox** 控件的优先级。对于固定命令栏，当一行中无法容纳其中的所有控件时，控件的优先级将决定该控件是否会从命令栏中删除。可读写。

**语法**

**express.Priority**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

有效的优先级数为从 0（零）至 7 之间的数字，默认值为 3。优先级为 1 表示不能从工具栏中删除该控件。而优先级为其他值的控件将被忽略。

菜单项的命令栏控件不能使用 **Priority** 属性。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Style**

获取或设置 **CommandBarComboBox** 控件的显示方式。该属性值可设置为以下 **MsoComboStyle** 常量之一：**msoComboLabel** 或 **msoComboNormal**。可读写。

**语法**

**express.Style**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Tag**

获取或设置有关 **CommandBarComboBox** 控件的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。

**语法**

**express.Tag**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Text**

获取或设置 **CommandBarComboBox** 控件的显示或编辑区中的文字。可读写。

**语法**

**express.Text**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例创建一个名为“Custom”的新命令栏，并向其添加一个包含有四个列表项的组合框。然后用 Text 属性将“Item 3”设置为默认列表项。  function test() {     let myBar = Application.CommandBars.Add("Custom", Application.Enum.msoBarTop, null, true)     myBar.Controls.Add(Application.Enum.msoControlComboBox, 1)     myBar.Visible = true     let testComboBox = Application.CommandBars.Item("Custom").Controls.Item(1)     testComboBox.Text = "Item 3"  }` |

#### **CommandBarComboBox.TooltipText**

获取或设置 **CommandBarComboBox**的**“屏幕提示”**中显示的文本。可读写。

**语法**

**express.TooltipText**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

默认情况下，**Caption** 属性的值将用作**“屏幕提示”**。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Top**

获取指定的 **CommandBarComboBox** 控件顶边到屏幕顶边的距离（以像素为单位）。只读。

**语法**

**express.Top**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Type**

获取 **CommandBarComboBox** 控件的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Visible**

获取或设置 **CommandBarComboBox** 控件的 **Visible** 属性。如果 **CommandBarControl** 可见，则此属性为 **True**。可读写。

**语法**

**express.Visible**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

#### **CommandBarComboBox.Width**

获取或设置指定的 **CommandBarComboBox** 控件的宽度（以像素为单位）。可读写。

**语法**

**express.Width**

*express*   一个代表 **CommandBarComboBox** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 在某些 WPS Office应用程序中，CommandBar 已被新的功能区用户界面取代。有关详细信息，请在帮助中搜索关键字“功能区”。 |

适用环境：web

适用平台：windows/linux