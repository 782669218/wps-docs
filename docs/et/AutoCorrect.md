**AutoCorrect**



包含 ET 的 AutoCorrect 属性（自动将日期名改为大写、自动更正连续两个大写字母、自动更正词条列表等等）。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用 AutoCorrect 属性可返回 AutoCorrect 对象。下例使 ET 更正以连续两个大写字母开头的单词。*/ function test() {     let correct = Application.AutoCorrect     correct.TwoInitialCapitals = true     correct.ReplaceText = true }` |

**方法**

|                                                              | 名称                  | 说明                                   |
| ------------------------------------------------------------ | --------------------- | -------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddReplacement**    | 向“自动更正”替换文本数组中添加项。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DeleteReplacement** | 删除“自动更正”替换数组中的一个输入项。 |

**属性**

|                                                              | 名称                          | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**               | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoExpandListRange**       | 表示列表的自动扩展是否被启用的 **Boolean** 值。在列表旁的空行或空列键入时，如果启用了自动扩展功能，则列表将扩展为包含此行或此列。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFillFormulasInLists**   | 影响由自动填充列表创建的计算列的创建。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CapitalizeNamesOfDays**     | 如果日期名称的第一个字母自动大写，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CorrectCapsLock**           | 如果 ET 自动更正无意中使用 Caps Lock 键，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CorrectSentenceCap**        | 如果 ET 自动更正句子（第一个单词）的大写，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                   | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayAutoCorrectOptions** | 允许用户显示或隐藏**“自动更正选项”**按钮。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                    | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReplaceText**               | 如果“自动更正”替换内容将被自动替换，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReplacementList**           | 返回“自动更正”替换内容的数组。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TwoInitialCapitals**        | 如果自动更正以两个大写字母开头的词，则该属性值为 **True**。**Boolean** 类型，可读写。 |

**成员方法**

#### **AutoCorrect.AddReplacement**

向“自动更正”替换文本数组中添加项。

**语法**

**express.AddReplacement(What, Replacement)**

*express*   一个代表 **AutoCorrect** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *What*        | 必选          | **String**   | 要替换的文本。如果该字符串已存在于自动更正替换文本数组中，则现有的替换文本将被新文本替换。 |
| *Replacement* | 必选          | **String**   | 替换文本。                                                   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在自动更正替换文字数组中，将单词“Paul”的替换单词指定为“Dreamer-Paul”。*/ Application.AutoCorrect.AddReplacement("Paul", "Dreamer-Paul")` |

#### **AutoCorrect.DeleteReplacement**

删除“自动更正”替换数组中的一个输入项。

**语法**

**express.DeleteReplacement(What)**

*express*   一个代表 **AutoCorrect** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *What*   | 必选          | **String**   | 要替换的文本（当文本出现在要从“自动更正”替换数组中删除的行上时）。如果该字符串不在“自动更正”替换数组中，此方法将失败。 |

**说明**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除“自动更正”替换数组中的“Paul”词条。*/ Application.AutoCorrect.DeleteReplacement("Paul")` |

**成员属性**

#### **AutoCorrect.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **AutoCorrect** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {   if(Application.AutoCorrect.Application.Value == "ET"){       alert("This is an ET Application object.")   }   else{       alert("This is not an ET Application object.")   } }` |

#### **AutoCorrect.AutoExpandListRange**

表示列表的自动扩展是否被启用的 **Boolean** 值。在列表旁的空行或空列键入时，如果启用了自动扩展功能，则列表将扩展为包含此行或此列。**Boolean** 类型，可读写。

**语法**

**express.AutoExpandListRange**

*express*   一个代表 **AutoCorrect** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例在相邻的行或列键入时启用了列表的自动扩展功能。*/ function test(){     Application.AutoCorrect.AutoExpandListRange = true } ` |

#### **AutoCorrect.AutoFillFormulasInLists**

影响由自动填充列表创建的计算列的创建。可读/写 **Boolean** 类型。

**语法**

**express.AutoFillFormulasInLists**

*express*   一个代表 **AutoCorrect** 对象的变量。

**说明**

该属性不影响现有计算列或非自动填充创建的计算列。

#### **AutoCorrect.CapitalizeNamesOfDays**

如果日期名称的第一个字母自动大写，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CapitalizeNamesOfDays**

*express*   一个代表 **AutoCorrect** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 ET 将日期名称的第一个字母大写。*/ function test() {   let correct = Application.AutoCorrect   correct.CapitalizeNamesOfDays = true   correct.ReplaceText = true }` |

#### **AutoCorrect.CorrectCapsLock**

如果 ET 自动更正无意中使用 Caps Lock 键，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CorrectCapsLock**

*express*   一个代表 **AutoCorrect** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 Caps Lock 的处理进行设置，使 ET 可自动更正无意中使用 Caps Lock。*/ Application.AutoCorrect.CorrectCapsLock = true` |

#### **AutoCorrect.CorrectSentenceCap**

如果 ET 自动更正句子（第一个单词）的大写，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CorrectSentenceCap**

*express*   一个代表 **AutoCorrect** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对大写功能的处理进行设置，使 ET 可自动更正句子大写。*/ Application.AutoCorrect.CorrectSentenceCap = true` |

#### **AutoCorrect.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **AutoCorrect** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **AutoCorrect.DisplayAutoCorrectOptions**

允许用户显示或隐藏**“自动更正选项”**按钮。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayAutoCorrectOptions**

*express*   一个代表 **AutoCorrect** 对象的变量。

**说明**

**DisplayAutoCorrectOptions** 属性是一个 WPS Office 范围的设置。在 ET 中更改该属性将影响其他的 Office 应用程序。

在 ET 中，只有在自动创建超链接时才会显示**“自动更正选项”**按钮。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定是否显示“自动更正选项”按钮并通知用户。*/ function CheckDisplaySetting(){     // Determine setting and notify user.     if(Application.AutoCorrect.DisplayAutoCorrectOptions == true){         alert("The AutoCorrect Options button can be displayed.")     }     else{         alert("The AutoCorrect Options button cannot be displayed.")     } }` |

#### **AutoCorrect.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **AutoCorrect** 对象的变量。

#### **AutoCorrect.ReplaceText**

如果“自动更正”替换内容将被自动替换，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ReplaceText**

*express*   一个代表 **AutoCorrect** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开自动文本替换功能。*/ function test() {   let correct = Application.AutoCorrect   correct.CapitalizeNamesOfDays = true   correct.ReplaceText = false }` |

#### **AutoCorrect.ReplacementList**

返回“自动更正”替换内容的数组。

**语法**

**express.ReplacementList**

*express*   一个代表 **AutoCorrect** 对象的变量。

**说明**

参数

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 可选          | **Variant**  | 要返回的“自动更正”替换内容数组的行索引。该行以二元一维数组的形式返回：第一个元素为第一列中的文本，第二个元素为第二列中的文本。 |

如果未指定 *Index*，则本方法将返回一个二维数组。数组中每一行包含一对自动更正替换文本，如下表所示。

| 列   | 内容           |
| ---- | -------------- |
| 1    | 要被替换的文本 |
| 2    | 替换文本       |

使用 [**AddReplacement**](https://qn.cache.wpscdn.cn/encs/doc/office_v13/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/AutoCorrect/%E6%96%B9%E6%B3%95/AddReplacement%20%E6%96%B9%E6%B3%95.htm) 方法可向替换列表添加自动更正项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在替换列表中搜索“Temperature”的替换条目，如果存在，则显示该替换条目。*/ function test() {   let repl = Application.AutoCorrect.ReplacementList      for(let i = 1; i <= repl.Count; i++){       if(repl(i,1) == "Temperature")       {         alert(i,2)       }   } }` |

#### **AutoCorrect.TwoInitialCapitals**

如果自动更正以两个大写字母开头的词，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.TwoInitialCapitals**

*express*   一个代表 **AutoCorrect** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 ET 自动更正以两个大写字母开头的词。*/ function test() {   let correct = Application.AutoCorrect   correct.TwoInitialCapitals = true   correct.ReplaceText = true }` |

适用环境：web

适用平台：windows/linux