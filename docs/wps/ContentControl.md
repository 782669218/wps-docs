**ContentControl** 



一个内容控件。内容控件是固定在文档中并可能带有标记的区域，这些区域作为特定类型的内容的容器。单个内容控件可能包含日期、列表或带格式文本段落等内容。**ContentControl** 对象是 **ContentControls** 集合的成员。

**说明**

使用 **ContentControls** 集合的 [Add ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ContentControls/ContentControls%20.htm#ContentControls.Add)方法可创建内容控件。使用 **Add** 方法的 *Type* 参数可以指定要创建的内容控件的型。以下示例创建了一个新的下拉列表内容控件，并在该列表中添加了若干个项。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)     //List entries     objCC.DropdownListEntries.Add("Cat")     objCC.DropdownListEntries.Add("Dog")     objCC.DropdownListEntries.Add("Horse")     objCC.DropdownListEntries.Add("Monkey")     objCC.DropdownListEntries.Add("Snake")     objCC.DropdownListEntries.Add("Other") }` |

使用 **Type** 属性可以更改内容控件的类型。例如，您或许希望将日期控件更改为文本控件。但是，您可能无法更改所有内容控件的类型；一些内容控件可能不允许更改其类型。此外，根据内容控件的内容，您也许不能更改控件的类型。例如，如果作为更改目标的内容控件不允许使用现有内容控件中内容的类型，那么，类型更改尝试将受到禁止并将生成运行时错误。

以下示例插入一个日期内容控件并且设置其值，然后将该控件更改为文本内容控件。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate)     objCC.Range.Text = ("January 1, 2007")     objCC.Type = wdContentControlText }` |

使用 **SetPlaceholderText** 方法可以将占位符文本从默认字符串更改为更适合该控件的内容。使用 **Title** 属性可以指定该控件的标题文本。在将光标放置到该控件的内部或者将鼠标指针放置到控件上方时，将在控件上方显示此标题。

根据您具有的内容控件的类型，您可能无法使用 **ContentControl** 对象的某些属性和方法。

并非所有内容控件属性都适用于所有不同类型的内容控件。下表列出了哪些属性适用于哪些类型的内容控件。

| 属性/方法                 | 适用于                                                       |
| ------------------------- | ------------------------------------------------------------ |
| BuildingBlockCategory属性 | BuildingBlock 库内容控件 (wdContentControlBuildingBlockGallery) |
| BuildingBlockType属性     | BuildingBlock 库内容控件 (wdContentControlBuildingBlockGallery) |
| DateDisplayFormat属性     | 日期内容控件 (wdContentControlDate)                          |
| DateDisplayLocale属性     | 日期内容控件 (wdContentControlDate)                          |
| DateStorageFormat属性     | 日期内容控件 (wdContentControlDate)                          |
| DropdownListEntries属性   | 组合框和下拉列表内容控件（wdContentControlComboBox 和 wdContentControlDropdownList） |
| MultiLine属性             | 纯文本内容控件 (wdContentControlText)                        |
| Ungroup方法               | 组内容控件 (wdContentControlGroup)                           |

**方法**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**               | 将活动文档中的内容控件复制到剪贴板。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**                | 从活动文档中删除内容控件并将该内容控件移动到剪贴板中。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**             | 删除指定的内容控件以及内容控件中的内容。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetCheckedSymbol**   | 设置用于代表复选框内容控件的选中状态的符号。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetPlaceholderText** | 设置在用户输入自己的文本之前显示在内容控件中的占位符文本。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetUncheckedSymbol** | 设置用于代表复选框内容控件的未选中状态的符号。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Ungroup**            | 从文档中删除一个组内容控件，以便它的子内容控件不再嵌套，并且可以自由地进行编辑。 |

**属性**

|                                                              | 名称                       | 说明                                                         |
| ------------------------------------------------------------ | -------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**            | 返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象代表 WPS 应用程序。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuildingBlockCategory**  | 返回或设置一个 **String**，代表生成块内容控件的类别。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuildingBlockType**      | 返回或设置一个 **WdBuildingBlockTypes** 常量，该常量代表生成块内容控件的生成块的类型。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Checked**                | 返回或设置 **Boolean** 类型的值，代表复选框的当前状态（选中/未选中）。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                | 返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DateCalendarType**       | 返回或设置一个 **WdCalendarType** 常量，该常量表示日历内容控件的日历类型。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DateDisplayFormat**      | 返回或设置 **String** 值，该值表示日期的显示格式。可读/写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DateDisplayLocale**      | 返回一个 **WdLanguageID**，它代表日期内容控件中显示的日期的语言格式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DateStorageFormat**      | 返回或设置 **WdContentControlDateStorageFormat** 值，该值表示在将日期内容控件绑定到活动文档的 XML 数据存储时的日期存储和检索格式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTextStyle**       | 返回或设置一个 **Variant**，它代表用于设置文本内容控件中文本格式的字符样式的名称。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DropdownListEntries**    | 返回 **ContentControlListEntries** 集合，该集合表示下拉列表内容控件或组合框内容控件中的项。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**                     | 返回一个代表内容控件标识的 **String** 类型的值。只读。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockContentControl**     | 返回或设置 **Boolean** 值，该值表示用户能否从活动文档中删除内容控件。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockContents**           | 返回或设置 **Boolean** 值，该值表示用户能否编辑内容控件的内容。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MultiLine**              | 返回或设置  **Boolean** 值，该值表示文本内容控件是否允许多行文本。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                 | 返回一个 **Object** 类型的值，该值代表指定 **ContentControl** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentContentControl**   | 返回一个 **ContentControl** 类型的值，该值代表嵌套在格式文本控件或组控件内的内容控件的父内容控件。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlaceholderText**        | 返回一个 **BuildingBlock** 对象，该对象代表内容控件的占位符文本。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**                  | 返回 **Range**，它表示活动文档中的内容控件的内容。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowingPlaceholderText** | 返回 **Boolean** 值，该值指示是否显示内容控件的占位符文本。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tag**                    | t返回或设置一个 **String** 类型的值，该值代表用于识别内容控件的值。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Temporary**              | 返回或设置 **Boolean** 值，该值表示当用户编辑内容控件的内容时是否从活动文档中删除该控件。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Title**                  | 返回或设置 **String** 值，该值表示内容控件的标题。可读/写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                   | 返回或设置 **WdContentControlType**，它表示内容控件的类型。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLMapping**             | 返回 **XMLMapping** 对象，该对象表示内容控件到文档的数据存储中的 XML 数据的映射。只读。 |

**成员方法**

#### **ContentControl.Copy**

将活动文档中的内容控件复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

使用 **Copy** 方法时，原始内容控件仍保留在活动文档中，但控件的副本（包括所有文本和属性设置）将移动到剪贴板中。然后，可以将内容控件粘贴到活动文档的其他部分。可以使用 **Selection** 对象的 **Paste** 方法或 **Range** 对象的 **Paste** 方法插入复制的内容控件，也可以使用 WPS 中的**“粘贴”**功能。

#### **ContentControl.Cut**

从活动文档中删除内容控件并将该内容控件移动到剪贴板中。

**语法**

**express.Cut()**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

可以使用 **Selection** 对象的 **Paste** 方法或 **Range** 对象的 **Paste** 方法插入剪切内容控件，也可以使用 WPS 中的**“粘贴”**功能。

#### **ContentControl.Delete**

删除指定的内容控件以及内容控件中的内容。

**语法**

**express.Delete(DeleteContents)**

*express*   一个代表 **ContentControl** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *DeleteContents* | 可选          | **Boolean**  | 指定是否删除内容控件的内容。如果值为 True，将同时删除内容控件及其内容。如果值为 False，将删除内容控件，但将其内容保留在活动文档中。默认值为 False。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 删除当前活动文档第一个内容控件 */ Application.ActiveDocument.ContentControls.Item(1).Delete() ` |

#### **ContentControl.SetCheckedSymbol**

设置用于代表复选框内容控件的选中状态的符号。

**语法**

**express.SetCheckedSymbol(CharacterNumber, Font)**

*express*   一个代表 **ContentControl** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *CharacterNumber* | 必选          | **Long**     | 指定符号的 Unicode 字符编号。该值始终是 31（字体开始处的控件符号的数目）加上此符号在符号表中的位置（从左向右计算）相应的数字。例如，若要指定一个位于 Symbol 字体符号表中第 37 号位置的 delta 字符，请将 CharacterNumber 设置为 68。 |
| *Font*            | 可选          | **String**   | 包含符号的字体的名称。                                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将指定内容控件的选中符号设置为“MS Gothic”字体“Ballot Box with X”符号。 */ function test() {     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlCheckBox)     objCC.SetCheckedSymbol(2612, "MS Gothic") }` |

#### **ContentControl.SetPlaceholderText**

设置在用户输入自己的文本之前显示在内容控件中的占位符文本。

**语法**

**express.SetPlaceholderText(BuildingBlock, Range, Text)**

*express*   一个代表 **ContentControl** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**      | **说明**                                              |
| --------------- | ------------- | ----------------- | ----------------------------------------------------- |
| *BuildingBlock* | 可选          | **BuildingBlock** | 指定 BuildingBlock 对象，该对象包含占位符文本的内容。 |
| *Range*         | 可选          | **Range**         | 指定 Range 对象，该对象包含占位符文本的内容。         |
| *Text*          | 可选          | **String**        | 指定占位符文本的内容。                                |

**说明**

在指定占位符文本时，只使用一个参数。如果使用多个参数，则 WPS 将使用在第一个参数中指定的文本。如果省略所有参数，则占位符文本将为空白。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中插入新的下拉列表内容控件，设置标题和占位符文本，然后在列表中插入几个新项。*/ function test() {     let objMap     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)     objCC.Title = "My Favorite Animal"     objCC.SetPlaceholderText, "Select your favorite animal "      /*List entries*/     objCC.DropdownListEntries.Add("Cat")     objCC.DropdownListEntries.Add("Dog")     objCC.DropdownListEntries.Add("Horse")     objCC.DropdownListEntries.Add("Monkey")     objCC.DropdownListEntries.Add("Snake")     objCC.DropdownListEntries.Add("Other") }` |

#### **ContentControl.SetUncheckedSymbol**

设置用于代表复选框内容控件的未选中状态的符号。

**语法**

**express.SetUncheckedSymbol(CharacterNumber, Font)**

*express*   一个代表 **ContentControl** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *CharacterNumber* | 必选          | **Long**     | 指定符号的 Unicode 字符编号。该值始终是 31（字体开始处的控件符号的数目）加上此符号在符号表中的位置（从左向右计算）相应的数字。例如，若要指定一个位于 Symbol 字体符号表中第 37 号位置的 delta 字符，请将 CharacterNumber 设置为 68。 |
| *Font*            | 可选          | **String**   | 包含符号的字体的名称。                                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将指定内容控件的未选中符号设置为“MS Gothic”字体“Ballot Box”符号。*/ function test() {     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlCheckBox)     objCC.SetUncheckedSymbol(2610, "MS Gothic") }` |

#### **ContentControl.Ungroup**

从文档中删除一个组内容控件，以便它的子内容控件不再嵌套，并且可以自由地进行编辑。

**语法**

**express.Ungroup()**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

如果您在非 **wdContentControlGroup** 类型的控件上运行该方法，则该方法将失败。

**成员属性**

#### **ContentControl.Application**

返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **ContentControl** 对象的变量。

#### **ContentControl.BuildingBlockCategory**

返回或设置一个 **String**，代表生成块内容控件的类别。可读写。

**语法**

**express.BuildingBlockCategory**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

此属性仅适用于生成块内容控件，并且相当于**“内容控件属性”**对话框中的**“类别”**选项。您可以将此属性设置为任何字符串；但如果您将它设置为不存在相应类别的字符串，则**“类别”**选项的值被设置为“(所有类别)”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例创建了一个新的生成块内容控件，并且指定了生成块的类型以及库。*/ function test() {     let objBB = Application.Selection.ContentControls.Add(wdContentControlBuildingBlockGallery)     objBB.BuildingBlockType = wdTypeEquations     objBB.BuildingBlockCategory = "General" }` |

#### **ContentControl.BuildingBlockType**

返回或设置一个 **WdBuildingBlockTypes** 常量，该常量代表生成块内容控件的生成块的类型。可读写。

**语法**

**express.BuildingBlockType**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

此属性仅适用于生成块内容控件，并且相当于**“内容控件属性”**对话框中的**“库”**选项。您只能为下列生成块类型设置此属性：

- 自定义 1 到自定义 5
- 自动图文集
- 文档部件
- 自定义自动图文集
- 自定义文档部件
- 公式

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例创建了一个新的生成块内容控件，并且指定了生成块的类型以及库。*/ function test() {     let objBB = Application.Selection.ContentControls.Add(wdContentControlBuildingBlockGallery)     objBB.BuildingBlockType = wdTypeEquations     objBB.BuildingBlockCategory = "General" }` |

#### **ContentControl.Checked**

返回或设置 **Boolean** 类型的值，代表复选框的当前状态（选中/未选中）。可读/写。

**语法**

**express.Checked**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

使用 **Checked** 属性获取/设置复选框内容控件的当前状态。如果该控件不是复选框，尝试访问该属性将失败，并导致运行时错误“该属性仅可用于复选框内容控件”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例设置指定复选框内容控件的 Checked 属性。*/ function test() {     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlCheckBox)     objCC.Title = "Send Reminder"     objCC.Checked = true }` |

#### **ContentControl.Creator**

返回一个 32 位的整数，该整数表示在其中创建加载项的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 中的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **ContentControl.DateCalendarType**

返回或设置一个 **WdCalendarType** 常量，该常量表示日历内容控件的日历类型。可读写。

**语法**

**express.DateCalendarType**

*express*   一个代表 **ContentControl** 对象的变量。

#### **ContentControl.DateDisplayFormat**

返回或设置 **String** 值，该值表示日期的显示格式。可读/写。

**语法**

**express.DateDisplayFormat**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

默认格式是在用户的系统上在 WPS 中指定的格式设置，通常取决于 Microsoft Windows 中的位置设置。例如，英语（美国）的默认日期格式为“mm/dd/yyyy”。使用 **DateDisplayFormat** 属性可以指定另一种日期格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate)     objCC.Title = "Review Period End Date"     objCC.DateDisplayFormat = "MMMM d, yyyy"     objCC.DateStorageFormat = wdContentControlDateStorageDate     objCC.Range.Text = "January 1, 2007" }` |

#### **ContentControl.DateDisplayLocale**

返回一个 **WdLanguageID**，它代表日期内容控件中显示的日期的语言格式。可读写。

**语法**

**express.DateDisplayLocale**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

默认情况下，**DateDisplayLocale** 属性最初设置为 WPS 中指定的编辑语言（前提是它与 Microsoft Windows 中指定的位置不同）。

#### **ContentControl.DateStorageFormat**

返回或设置 **WdContentControlDateStorageFormat** 值，该值表示在将日期内容控件绑定到活动文档的 XML 数据存储时的日期存储和检索格式。可读/写。

**语法**

**express.DateStorageFormat**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

**DateStorageFormat** 属性用于以日期格式、日期/时间格式或文本格式存储日期。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将日期内容控件添加到活动文档中，并指定日期、日期显示格式和日期存储格式。*/ function test() {     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate)     objCC.Title = "Review Period End Date"     objCC.DateDisplayFormat = "MMMM d, yyyy"     objCC.DateStorageFormat = wdContentControlDateStorageDate     objCC.Range.Text = "January 1, 2007" }` |

#### **ContentControl.DefaultTextStyle**

返回或设置一个 **Variant**，它代表用于设置文本内容控件中文本格式的字符样式的名称。可读写。

**语法**

**express.DefaultTextStyle**

*express*   一个代表 **ContentControl** 对象的变量。

#### **ContentControl.DropdownListEntries**

返回 **ContentControlListEntries** 集合，该集合表示下拉列表内容控件或组合框内容控件中的项。只读。

**语法**

**express.DropdownListEntries**

*express*   一个代表 **ContentControl** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中插入新的下拉列表内容控件，设置标题和占位符文本，然后向列表中添加几个新项。*/ function test() {     let objMap     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)     objCC.Title = "My Favorite Animal"     objCC.SetPlaceholderText, "Select your favorite animal "      /*List entries*/     objCC.DropdownListEntries.Add("Cat")     objCC.DropdownListEntries.Add("Dog")     objCC.DropdownListEntries.Add("Horse")     objCC.DropdownListEntries.Add("Monkey")     objCC.DropdownListEntries.Add("Snake")     objCC.DropdownListEntries.Add("Other") }` |

#### **ContentControl.ID**

返回一个代表内容控件标识的 **String** 类型的值。只读。

**语法**

**express.ID**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

**ID** 属性是一个内部号码，您不能更改它，但可以使用它在代码中标识内容控件。此号码对于每个内容控件都是唯一的，并且不会更改。

**示例**

| 示例代码复制                                            |
| ------------------------------------------------------- |
| `Application.ActiveDocument.ContentControls.Item(1).ID` |

#### **ContentControl.LockContentControl**

返回或设置 **Boolean** 值，该值表示用户能否从活动文档中删除内容控件。可读/写。

**语法**

**express.LockContentControl**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

此属性的默认值为 **False**。此属性对应于**“内容控制属性”**对话框中的**“不能删除内容控件”**复选框。如果 Temporary属性设置为 **True**，则您不能设置此属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 插入日期内容控件，设置内容后，并且禁止编辑以及禁止删除 */ function test() { 	let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate) 	objCC.Range.Text = "January 1, 2007" 	objCC.LockContents = true 	objCC.LockContentControl = true }` |

#### **ContentControl.LockContents**

返回或设置 **Boolean** 值，该值表示用户能否编辑内容控件的内容。可读/写。

**语法**

**express.LockContents**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

此属性的默认值为 **False**。此属性对应于**“内容控制属性”**对话框中的**“不能编辑内容”** 复选框。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 设置当前文档第一个内容控件可以编辑 */ Application.ActiveDocument.ContentControls.Item(1).LockContents = false` |

#### **ContentControl.MultiLine**

返回或设置  **Boolean** 值，该值表示文本内容控件是否允许多行文本。可读/写。

**语法**

**express.MultiLine**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

此属性只应用于文本内容控件和格式文本内容控件。

#### **ContentControl.Parent**

返回一个 **Object** 类型的值，该值代表指定 **ContentControl** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ContentControl** 对象的变量。

#### **ContentControl.ParentContentControl**

返回一个 **ContentControl** 类型的值，该值代表嵌套在格式文本控件或组控件内的内容控件的父内容控件。只读。

**语法**

**express.ParentContentControl**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

如果内容控件未嵌套，则此属性返回错误。

#### **ContentControl.PlaceholderText**

返回一个 **BuildingBlock** 对象，该对象代表内容控件的占位符文本。只读。

**语法**

**express.PlaceholderText**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

可以使用 **SetPlaceholderText** 方法设置内容控件的占位符文本。

#### **ContentControl.Range**

返回 **Range**，它表示活动文档中的内容控件的内容。只读。

**语法**

**express.Range**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

使用 **Range** 属性可以访问文本、文本格式和其他文本属性。有关详细信息，请参阅 处理 Range 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 插入日期内容控件，设置内容 */ function test() { 	let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDate) 	objCC.Range.Text = "January 1, 2007" }` |

#### **ContentControl.ShowingPlaceholderText**

返回 **Boolean** 值，该值指示是否显示内容控件的占位符文本。只读。

**语法**

**express.ShowingPlaceholderText**

*express*   一个代表 **ContentControl** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 添加一个文本内容控件，如果可以设置占位文本，那就设置为"hello" */ function test() { 	let txt = Application.ActiveDocument.ContentControls.Add(wdContentControlText); 	if (txt.ShowingPlaceholderText) { 		txt.SetPlaceholderText(null, null, "hello"); 	} }` |

#### **ContentControl.Tag**

t返回或设置一个 **String** 类型的值，该值代表用于识别内容控件的值。可读写。

**语法**

**express.Tag**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

**Tag** 属性与 **Title** 属性的不同之处在于，用户在编辑文档时从不会显示标记。当文档处于打开状态时，开发人员可以利用该属性来存储编程操作的值。

#### **ContentControl.Temporary**

返回或设置 **Boolean** 值，该值表示当用户编辑内容控件的内容时是否从活动文档中删除该控件。可读/写。

**语法**

**express.Temporary**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

默认值为 **False**。此属性对应于**“内容控件属性”** 对话框中的**“内容被编辑后删除内容控件”**复选框。

| 注释                                                         |
| ------------------------------------------------------------ |
| 如果 **LockContentControl** 属性设置为 **True**，则您不能设置此属性。 |

#### **ContentControl.Title**

返回或设置 **String** 值，该值表示内容控件的标题。可读/写。

**语法**

**express.Title**

*express*   一个代表 **ContentControl** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 插入文本内容控件并设置标题 */ function test() { 	let txt = Application.ActiveDocument.ContentControls.Add(wdContentControlText); 	txt.Title = "txt title" }` |

#### **ContentControl.Type**

返回或设置 **WdContentControlType**，它表示内容控件的类型。可读/写。

**语法**

**express.Type**

*express*   一个代表 **ContentControl** 对象的变量。

**说明**

可以使用 **Type** 属性将内容控件的类型从一种类型更改为另一种类型。但是，能否更改控件的类型将取决于原始类型和更改时内容控件中的内容。所有内容控件都可以更改为格式文本或生成块库类型的内容控件，因为这些类型允许任意内容。对于其他类型，如果内容对于要更改为的目标类型来说是有效的，则允许更改类型。否则，将拒绝更改，并导致运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例检查并确定指定的内容控件是否是下拉列表框或组合框。如果它是这两种类型之一，则上移列表的最后一项，使它成为列表的第一项。 */ function test() {     let objCC = Application.ActiveDocument.ContentControls.Item(1);     if (objCC.Type == wdContentControlComboBox || objCC.Type == wdContentControlDropdownList) {         let objCL = objCC.DropdownListEntries.Item(objCC.DropdownListEntries.Count)         for (let i = 1; i <= objCC.DropdownListEntries.Count - 1; i++) {             objCL.MoveUp()         }     } } ` |

#### **ContentControl.XMLMapping**

返回 **XMLMapping** 对象，该对象表示内容控件到文档的数据存储中的 XML 数据的映射。只读。

**语法**

**express.XMLMapping**

*express*   一个代表 **ContentControl** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置内置的“作者”文档属性，并向活动文档添加新内容控件，然后设置该控件到该内置文档属性的值的映射。*/ function test() {     Application.ActiveDocument.BuiltInDocumentProperties.Item("Author").Value = "David Jaffe"     let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlText, Application.ActiveDocument.Paragraphs.Item(1).Range)     let objMap = objCC.XMLMapping     blnMap = objMap.SetMapping("/ns1:coreProperties[1]/ns0:creator[1]")     if (blnMap == false) {         alert("Unable to map the content control.")     } }` |

适用环境：web

适用平台：windows/linux