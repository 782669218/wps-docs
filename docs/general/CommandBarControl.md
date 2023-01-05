#### **CommandBarControl**



代表一个命令栏控件。**CommandBarControl** 对象是 **CommandBarControls** 集合的成员。**CommandBarControl** 对象与 **CommandBarButton**、**CommandBarComboBox** 和 **CommandBarPopup** 对象具有同样的属性和方法。

**说明**

在编写处理自定义命令栏控件的 Visual Basic 代码时，请使用 **CommandBarButton**、**CommandBarComboBox** 和 **CommandBarPopup** 对象。在编写用于处理容器应用程序中的内置控件的代码时，如果该控件不能用上述三个对象中的任意一个来代表，则可以使用 **CommandBarControl** 对象。使用 **Controls**(*index*) 可返回一个 **CommandBarControl** 对象，其中 *index* 是控件的索引号。（该控件的 **Type** 属性必须是 **msoControlLabel**、**msoControlExpandingGrid**、**msoControlSplitExpandingGrid**、**msoControlGrid** 或 **msoControlGauge**）。声明为 **CommandBarControl** 的变量可赋予 **CommandBarButton**、**CommandBarComboBox** 和 **CommandBarPopup** 值

**方法**

|                                                              | 名称        | 说明                                                        |
| ------------------------------------------------------------ | ----------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**    | 将一个命令栏控件复制到已有的命令栏中。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**  | 将 **CommandBarControl** 对象从其集合中删除。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Execute** | 运行分配给指定 **CommandBarControl** 控件的过程或内置命令。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**    | 将指定的 **CommandBarControl** 移动到已有的命令栏。         |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**         | 获取或设置命令栏控件的标题文字。可读写。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DescriptionText** | 获取或设置命令栏控件的说明。可读写。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enabled**         | 获取或设置指定是否启用 **CommandBarControl** 的 **Boolean** 值。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**              | 获取内置的 **CommandBarControl** 的 ID。只读。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tag**             | 获取或设置有关 **CommandBarControl** 的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visibl**          | 获取或设置 **CommandBarControl** 的 **Visible** 属性。如果 **CommandBarControl** 可见，则此属性为 **true**。可读写。 |

**成员方法**

#### **CommandBarControl.Copy**

将一个命令栏控件复制到已有的命令栏中。

**语法**

**express.Copy(Bar, Before)**

*express*   一个代表 **CommandBarControl** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Bar*    | 可选          | **Variant**  | 一个代表目标命令栏的 CommandBar 对象。如果省略此参数，控件将被复制到它原来所在的命令栏中。 |
| *Before* | 可选          | **Variant**  | 一个指示新控件在命令栏上的位置的数值。新控件将插入到位于此位置的控件之前。如果省略此参数，控件将被复制到命令栏的末尾。 |

#### **CommandBarControl.Delete**

将 **CommandBarControl** 对象从其集合中删除。

**语法**

**express.Delete(Temporary)**

*express*   一个代表 **CommandBarControl** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Temporary* | 可选          | **Variant**  | 如果为 true，则从当前会话中删除此控件。应用程序在下次会话中将再次显示该控件 |

#### **CommandBarControl.Execute**

运行分配给指定 **CommandBarControl** 控件的过程或内置命令。

**语法**

**express.Execute()**

*express*   一个代表 **CommandBarControl** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本 ET 示例创建一个命令栏，然后向其中添加内置命令栏按钮控件。该按钮将执行 ET AutoSum 函数。本示例使用 Execute 方法在显示该命令栏时计算选定单元格区域的总计。 */  function test() { 	let cmdBars = Application.ActiveDocument.CommandBars 	let cbrCustBar = cmdBars.Add("Custom") 	let ctlAutoSum = cbrCustBar.Controls.Add(Application.Enum.msoControlButton, cmdBars.Item("Standard").Controls.Item("AutoSum").Id) 	cbrCustBar.Visible = true  	ctlAutoSum.Execute() }` |

#### **CommandBarControl.Move**

将指定的 **CommandBarControl** 移动到已有的命令栏。

**语法**

**express.Move(Bar, Before)**

*express*   一个代表 **CommandBarControl** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Bar*    | 可选          | **Variant**  | 一个代表控件的目标命令栏的 Command 对象。如果忽略该参数，则控件将移动到当前所在命令栏的末端。 |
| *Before* | 可选          | **Variant**  | 表示控件位置的数字。控件将插到该位置的控件之前。如果忽略该参数，控件插入到同一命令栏。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将名为 Custom 的命令栏上的第一个组合框控件移动到该命令栏上的第七个控件之前。本示例将 Tag 设置为“Selection box”并赋予控件较低的优先级，以便在一行容纳不下所有控件时将其隐藏。 */ function test { 	let allcontrols = Application.ActiveDocument.CommandBars.Item("Custom").Controls 	for(let i = 1; i <= allControls.Count; i++){     if(allControls.Item(i).Type == Application.Enum.msoControlComboBox){         let newallcontrols = allControls.Item(i)             newallcontrols.Move(null, 7)             newallcontrols.Tag = "Selection box"             newallcontrols.Priority = 5         break;     } }	 }` |

**成员属性**

#### **CommandBarControl.Caption**

获取或设置命令栏控件的标题文字。可读写。

**语法**

**express.Caption**

*express*   一个代表 **CommandBarControl** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在自定义命令栏中添加一个带拼写检查按钮图标的命令栏控件，然后将其标题设置为“Spelling checker”。 */  function test() { 	let myBar = Application.ActiveDocument.CommandBars.Add("Custom", Application.Enum.msoBarTop, true) 	myBar.Visible = true  	let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, 2) 	myControl.DescriptionText = "Starts the spelling checker"  	myControl.Caption = "Spelling checker" }` |

#### **CommandBarControl.DescriptionText**

获取或设置命令栏控件的说明。可读写。

**语法**

**express.DescriptionText**

*express*   一个代表 **CommandBarControl** 对象的变量。

**说明**

该说明不显示给用户，但有助于其他开发者将控件行为归档。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在自定义命令栏中添加一个带控件行为说明的控件。 */ function test() { 	let commandbars = Application.ActiveDocument.CommandBars 	let myBar = commandbars.Add("Custom", Application.Enum.msoBarTop, true) 	myBar.Visible = true  	let myControl = myBar.Controls.Add(Application.Enum.msoControlButton, commandbars.Item("Standard").Controls.Item("Paste").Id) 	myControl.DescriptionText = "Pastes the contents of the Clipboard"         myControl.Caption = "Paste" }` |

#### **CommandBarControl.Enabled**

获取或设置指定是否启用 **CommandBarControl** 的 **Boolean** 值。可读写。

**语法**

**express.Enabled**

*express*   一个代表 **CommandBarControl** 对象的变量。

**说明**

对于命令栏，如果将该属性设置为 **True**，则该命令栏的名称将出现在可用命令栏列表中。

对于内置控件，如果将 **Enabled** 属性设置为 **True**，则将由应用程序确定其状态。但如果将该属性设置为 **False**，则强行禁用该控件。

#### **CommandBarControl.Id**

获取内置的 **CommandBarControl** 的 ID。只读。

**语法**

**express.Id**

*express*   一个代表 **CommandBarControl** 对象的变量。

#### **CommandBarControl.Tag**

获取或设置有关 **CommandBarControl** 的信息，例如，可作为过程参数的数据或用于识别该控件的信息。可读写。

**语法**

**express.Tag**

*express*   一个代表 **CommandBarControl** 对象的变量。

#### **CommandBarControl.Visibl**

获取或设置 **CommandBarControl** 的 **Visible** 属性。如果 **CommandBarControl** 可见，则此属性为 **true**。可读写。

**语法**

**express.Visibl**

*express*   一个代表 **CommandBarControl** 对象的变量。

适用环境：web

适用平台：windows/linux