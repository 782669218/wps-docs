**Workbooks**



ET 应用程序中当前打开的所有 [Workbook ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbooks/Workbooks%20.htm#jsObject_Workbooks)对象的集合。

**说明**

有关使用一个 **Workbook** 对象的详细信息，请参阅 [Workbook](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象。

## 示例

使用 [Workbooks ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbooks/Workbooks%20.htm#jsObject_Workbooks)属性可返回 [Workbooks](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbooks/Workbooks%20.htm#jsObject_Workbooks) 集合。下例关闭所有打开的工作簿。

| 示例代码复制                    |
| ------------------------------- |
| `Application.Workbooks.Close()` |

使用 [Add ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbooks/Workbooks%20.htm#Workbooks.Add)方法可创建一个新空工作簿并将它添加到集合。下例给 ET 添加一个新空工作簿。

| 示例代码复制                  |
| ----------------------------- |
| `Application.Workbooks.Add()` |

使用 [Open ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbooks/Workbooks%20.htm#Workbooks.Open)方法可打开一个文件。这样会为打开的文件创建一个新工作簿。下例将文件 Array.xls 打开为只读工作簿。

| 示例代码复制                                        |
| --------------------------------------------------- |
| `Application.Workbooks.Open("Array.xls",null,true)` |

**方法**

|                                                              | 名称         | 说明                                                         |
| ------------------------------------------------------------ | ------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**      | 新建一个工作表。新工作表将成为活动工作表。返回值：一个代表新工作簿的 [Workbook ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Close**    | 关闭对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**     | 从集合中返回一个对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Open**     | 打开一个工作簿。返回值：一个代表打开的工作簿的 [Workbook ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OpenText** | 载入一个文本文件，并将其作为包含单个工作表的新工作簿进行分列处理，然后在此工作表中放入经过分列处理的文本文件数据。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |

**成员方法**

#### **Workbooks.Add**

新建一个工作表。新工作表将成为活动工作表。

返回值：一个代表新工作簿的 [Workbook ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象。

**语法**

**express.Add(Template)**

*express*   一个代表 **Workbooks** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Template* | 可选          | **Variant**  | 确定如何创建新工作簿。如果此参数为指定现有 ET 文件名的字符串，那么创建新工作簿将以该指定的文件作为模板。如果此参数为常量，新工作簿将包含一个指定类型的工作表。可为以下 XlWBATemplate 常量之一：xlWBATChart、xlWBATExcel4IntlMacroSheet、xlWBATExcel4MacroSheet 或 xlWBATWorksheet。如果省略此参数，ET 将创建包含一定数目空白工作表的新工作簿（该数目由 SheetsInNewWorkbook 属性设置）。 |

**说明**

如果 *Template* 参数指定的是文件，则该文件名可包含路径。

#### **Workbooks.Close**

关闭对象。

**语法**

**express.Close()**

*express*   一个代表 **Workbooks** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例关闭所有打开的工作簿。如果某个打开的工作簿有改动，ET 将显示询问是否保存更改的对话框和相应提示。*/ Application.Workbooks.Close()` |

#### **Workbooks.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Workbooks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将变量 wb 设置为 Myaddin.xla 的工作簿。*/ let wb = Application.Workbooks.Item("myaddin.xla")` |

#### **Workbooks.Open**

打开一个工作簿。返回值：一个代表打开的工作簿的 [Workbook ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象。

**语法**

**express.Open(FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad)**

*express*   一个代表 **Workbooks** 对象的变量。

**参数**

| **名称**                    | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*                  | 可选          | **Variant**  | String. 要打开的工作簿的文件名。                             |
| *UpdateLinks*               | 可选          | **Variant**  | 指定更新文件中外部引用（链接）的方式，如下面的公式 =SUM([Budget.xls]Annual!C10:C25) 中对 Budget.xls 工作簿中某个区域的引用。如果省略此参数，则提示用户指定链接的更新方式。有关此参数所用值的详细信息，请参阅“说明”部分。如果 ET 正在打开 WKS、WK1 或 WK3 格式的文件，并且 UpdateLinks 参数为 0，则不创建任何图表；否则 ET 根据附加于该文件的图形生成图表。 |
| *ReadOnly*                  | 可选          | **Variant**  | 如果为 True，则以只读模式打开工作簿。                        |
| *Format*                    | 可选          | **Variant**  | 如果 ET 正在打开文本文件，则由此参数指定分隔符。如果省略此参数，则使用当前的分隔符。有关此参数值的详细信息，请参阅“备注”部分。 |
| *Password*                  | 可选          | **Variant**  | 一个字符串，包含打开受保护工作簿所需的密码。如果省略此参数并且工作簿已设置密码，则提示用户输入密码。 |
| *WriteResPassword*          | 可选          | **Variant**  | {description} 一个字符串，包含写入受保护工作簿所需的密码。如果省略此参数并且工作簿已设置密码，则提示用户输入密码。 |
| *IgnoreReadOnlyRecommended* | 可选          | **Variant**  | 如果为 True，则不让 ET 显示只读的建议消息（如果该工作簿以“建议只读”选项保存）。 |
| *Origin*                    | 可选          | **Variant**  | 如果该文件为文本文件，则此参数用于指示该文件来源于何种操作系统（以便正确映射代码页和回车/换行符 (CR/LF)）。可为以下 XlPlatform 常量之一：xlMacintosh、xlWindows 或 xlMSDOS。如果省略此参数，则使用当前操作系统。 |
| *Delimiter*                 | 可选          | **Variant**  | 如果该文件为文本文件并且 Format 参数为 6，则此参数是一个字符串，指定用作分隔符的字符。例如，可使用 Chr(9) 代表制表符，使用“,”代表逗号，使用“;”代表分号，或者使用自定义字符。只使用字符串的第一个字符。 |
| *Editable*                  | 可选          | **Variant**  | 如果文件为 ET 4.0 加载宏，则此参数为 True 时可打开该加载宏以使其在窗口中可见。如果此参数为 False 或被省略，则以隐藏方式打开加载宏，并且无法设为可见。本选项不能应用于由 ET 5.0 或更高版本的 ET 创建的加载宏。如果文件是 ET 模板，则参数值为 True 时，会打开指定模板进行编辑。参数值为 False 时，可根据指定模板打开新的工作簿。默认值为 False。 |
| *Notify*                    | 可选          | **Variant**  | 当文件不能以可读写模式打开时，如果此参数为 True，则可将该文件添加到文件通知列表。ET 将以只读模式打开该文件并轮询文件通知列表，并在文件可用时向用户发出通知。如果此参数为 False 或被省略，则不请求任何通知，并且不能打开任何不可用的文件。 |
| *Converter*                 | 可选          | **Variant**  | 打开文件时试用的第一个文件转换器的索引。首先试用的是指定的文件转换器；如果该转换器不能识别此文件，则试用所有其他转换器。转换器索引由 FileConverters 属性返回的转换器行号组成。 |
| *AddToMru*                  | 可选          | **Variant**  | 如果为 True，则将该工作簿添加到最近使用的文件列表中。默认值为 False。 |
| *Local*                     | 可选          | **Variant**  | 如果为 True，则以 ET（包括控制面板设置）的语言保存文件。如果为 False（默认值），则以 示例代码 (VBA) 语言保存文件。VBA 通常为美国英语版本，除非从中运行 Workbooks.Open 的 VBA 项目是旧的国际化 XL5/95 VBA 项目。 |
| *CorruptLoad*               | 可选          | **Variant**  | 可为以下常量之一：xlNormalLoad、xlRepairFile 和 xlExtractData。如果未指定任何值，则默认行为是 xlNormalLoad，并且当通过 OM 启动时不尝试恢复状态。 |

**说明**

默认情况下，以编程方式打开文件时将启用宏。使用 [AutomationSecurity ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/MsoAutomationSecurity%20%E6%9E%9A%E4%B8%BE.html)属性可设置以编程方式打开文件时所用的宏安全模式。

可在 *UpdateLinks* 参数中指定下面的一个值，以确定在工作簿打开时是否更新外部引用（链接）：

| 值   | 含义                                 |
| ---- | ------------------------------------ |
| 0    | 工作簿打开时不更新外部引用（链接）。 |
| 3    | 工作簿打开时更新外部引用（链接）。   |

您可在 *Format* 参数中指定下面的一个值，以确定文件的分隔字符：

| Value | 分隔符                              |
| ----- | ----------------------------------- |
| 1     | 标签                                |
| 2     | 逗号                                |
| 3     | 空格                                |
| 4     | 分号                                |
| 5     | Nothing                             |
| 6     | 自定义字符（请参阅 Delimiter 参数） |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开 Analysis.xls 工作簿，然后运行它的 Auto_Open 宏。*/ function test() {     Application.Workbooks.Open("ANALYSIS.XLS")     Application.ActiveWorkbook.RunAutoMacros(xlAutoOpen) }` |

#### **Workbooks.OpenText**

载入一个文本文件，并将其作为包含单个工作表的新工作簿进行分列处理，然后在此工作表中放入经过分列处理的文本文件数据。

**语法**

**express.OpenText(Filename, Origin, StartRow, DataType, TextQualifier, ConsecutiveDelimiter, Tab, Semicolon, Comma, Space, Other, OtherChar, FieldInfo, TextVisualLayout, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers, Local)**

*express*   一个代表 **Workbooks** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型**        | **说明**                                                     |
| ---------------------- | ------------- | ------------------- | ------------------------------------------------------------ |
| *Filename*             | 必选          | **String**          | 指定要打开和分列的文本文件的名称。                           |
| *Origin*               | 可选          | **Variant**         | 指定文本文件来源。可为以下 XlPlatform 常量之一：xlMacintosh、xlWindows 或 xlMSDOS。此外，它还可以是一个整数，表示所需代码页的代码页编号。例如，“1256”指定源文本文件的编码是阿拉伯语 (Windows)。如果省略该参数，则此方法将使用“文本导入向导”中“文件原始格式”选项的当前设置。 |
| *StartRow*             | 可选          | **Variant**         | 文本分列处理的起始行号。默认值为 1。                         |
| *DataType*             | 可选          | **Variant**         | 指定文件中数据的列格式。可为以下 XlTextParsingType 常量之一：xlDelimited 或 xlFixedWidth。如果未指定该参数，则 ET 将尝试在打开文件时确定列格式。 |
| *TextQualifier*        | 可选          | **XlTextQualifier** | 指定文本识别符号。                                           |
| *ConsecutiveDelimiter* | 可选          | **Variant**         | 如果为 True，则将连续分隔符视为一个分隔符。默认值为 False。  |
| *Tab*                  | 可选          | **Variant**         | 如果为 True，则将制表符用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。 |
| *Semicolon*            | 可选          | **Variant**         | 如果为 True，则将分号用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。 |
| *Comma*                | 可选          | **Variant**         | 如果为 True，则将逗号用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。 |
| *Space*                | 可选          | **Variant**         | 如果为 True，则将空格用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。 |
| *Other*                | 可选          | **Variant**         | 如果为 True，则将 OtherChar 参数指定的字符用作分隔符（DataType 必须为 xlDelimited）。默认值为 False。 |
| *OtherChar*            | 可选          | **Variant**         | （如果 Other 为 True，则为必选项）。当 Other 为 True 时，指定分隔符。如果指定了多个字符，则仅使用字符串中的第一个字符而忽略剩余字符。 |
| *FieldInfo*            | 可选          | **Variant**         | 包含单列数据相关分列信息的数组。对该参数的解释取决于 DataType 的值。如果此数据由分隔符分隔，则该参数为由两元素数组组成的数组，其中每个两元素数组指定一个特定列的转换选项。第一个元素为列标（从 1 开始），第二个元素是 XlColumnDataType 的常量之一，用于指定分列方式。 |
| *TextVisualLayout*     | 可选          | **Variant**         | 文本的可视布局。                                             |
| *DecimalSeparator*     | 可选          | **Variant**         | 识别数字时，ET 使用的小数分隔符。默认设置为系统设置。        |
| *ThousandsSeparator*   | 可选          | **Variant**         | 识别数字时， ET 使用的千位分隔符。默认设置为系统设置。       |
| *TrailingMinusNumbers* | 可选          | **Variant**         | 如果应将结尾为减号字符的数字视为负数处理，则指定为 True。如果为 False 或省略该参数，则将结尾为减号字符的数字视为文本处理。 |
| *Local*                | 可选          | **Variant**         | 如果分隔符、数字和数据格式应使用计算机的区域设置，则指定为 True。 |

**说明**

**FieldInfo 参数信息**

只有在安装并选定了中国台湾地区语言支持时才可使用 **xlEMDFormat**。**xlEMDFormat** 常量指定使用中国台湾地区纪元日期。

列说明符可为任意顺序。输入数据中如果某列没有列说明符，则用常规设置对该列进行分列处理。

**注释：**

- 在指定要跳过的列时，必须明确说明其余各列的类型，否则数据不能正确拆分。
- 如果数据中包含可识别的日期，则工作表中单元格的格式就会设置为“日期”，即便为该列设置的格式为“常规”。此外，如果您为某一列指定了以上日期格式中的一种格式，而数据中并不包含可识别的日期，那么工作表中单元格的格式为“常规”。

如果源数据为定宽列，则每个两元数组的第一个元素指定起始元素在列中的位置，（用整数表示，第一个字符为 0 (零)），第二个元素用 0 到 9 的数字指定分列选项，如前表所示。

**ThousandsSeparator 参数信息**

下表显示了使用不同的导入设置向 ET 中导入文本时的结果。数字结果显示在最右边的列中。

| 系统小数分隔符 | 系统千位分隔符 | 小数分隔符值 | 千位分隔符值 | 导入的文本 | 单元格的值（数据类型） |
| -------------- | -------------- | ------------ | ------------ | ---------- | ---------------------- |
| 句号           | 逗号           | 逗号         | 句号         | 123.123,45 | 123,123.45（数字）     |
| 句号           | 逗号           | 逗号         | 逗号         | 123.123,45 | 123.123,45（文本）     |
| 逗号           | 句号           | 句号         | 逗号         | 123,123.45 | 123,123.45（数字）     |
| 句号           | 逗号           | 句号         | 逗号         | 123 123.45 | 123 123.45（文本）     |
| 句号           | 逗号           | 句号         | 空格         | 123 123.45 | 123,123.45（数字）     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开 Data.txt 文件并将制表符作为分隔符对此文件进行分列处理，将其转换成为工作表。*/ Application.Workbooks.OpenText("C:\\DATA.TXT",null,null,xlDelimited,null,null,true)` |

**成员属性**

#### **Workbooks.Application**

如果不使用对象识别符，则该属性返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Workbooks** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }else {         alert("This is not an ET Application object.")     } }` |

#### **Workbooks.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Workbooks** 对象的变量。

适用环境：web

适用平台：windows/linux