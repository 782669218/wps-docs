**Sheets**



指定的或活动工作簿中所有工作表的集合。

**说明**

**Sheets** 集合可以包含 **Chart** 或 **Worksheet**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Worksheet/Worksheet%20.htm#jsObject_Worksheet)对象。

如果希望返回所有类型的工作表，[**Sheets** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#jsObject_Sheets)集合就非常有用。如果仅需使用某一类型的工作表，请参阅该工作表类型的对象主题。

使用 **Sheets** 属性可返回 **Sheets** 集合。下例打印当前活动工作簿上的所有工作表。

| 示例代码复制                    |
| ------------------------------- |
| `Application.Sheets.PrintOut()` |

使用 **Add**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Sheets/Sheets%20.htm#Sheets.Add)方法可创建一个新的工作表并将它添加到集合。下例给活动工作簿添加两个图表工作表，将它们放在工作簿中的工作表二之后。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Sheets.Add(null, Application.Sheets.Item(2),2,xlChart)` |

使用 **Sheets**(*index*)（其中 *index* 是工作表名称或索引号）可返回一个 **Chart** 或 **Worksheet**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Worksheet/Worksheet%20.htm#jsObject_Worksheet)对象。下例激活名为“Sheet1”的工作表。

| 示例代码复制                                   |
| ---------------------------------------------- |
| `Application.Sheets.Item("sheet1").Activate()` |

使用 **Sheets**(*array*) 可指定多个工作表。下例将名为“Sheet4”和“Sheet5”的工作表移到工作簿的开头。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Sheets.Item(["Sheet4", "Sheet5"]).Move(Application.Sheets.Item(1))` |

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**              | 新建工作表、图表或宏表。新建的工作表将成为活动工作表。返回值：一个 Object 值，它代表新的工作表、图表或宏表。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**             | 将工作表复制到工作簿的另一位置。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**           | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FillAcrossSheets** | 将单元格区域复制到集合中所有其他工作表的同一位置。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**             | 从集合中返回一个对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**             | 将工作表移到工作簿中的其他位置。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintOut**         | 打印对象。返回Variant值                                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintPreview**     | 按对象打印后的外观效果显示对象的预览。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**           | 选择对象。                                                   |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HPageBreaks** | 返回一个 **HPageBreaks**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/HPageBreaks/HPageBreaks%20.htm#jsObject_HPageBreaks)集合，它代表工作表上的水平分页符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VPageBreaks** | 返回一个 **VPageBreaks**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/VPageBreaks/VPageBreaks%20.htm#jsObject_VPageBreaks)集合，它代表工作表上的垂直分页符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**     | 返回或设置一个 **Variant** 值，它确定对象是否可见。          |

**成员方法**

#### **Sheets.Add**

新建工作表、图表或宏表。新建的工作表将成为活动工作表。

返回值：一个 Object 值，它代表新的工作表、图表或宏表。

**语法**

**express.Add(Before, After, Count, Type)**

*express*   一个代表 **Sheets** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before* | 可选          | **Variant**  | 指定工作表的对象，新建的工作表将置于此工作表之前。           |
| *After*  | 可选          | **Variant**  | 指定工作表的对象，新建的工作表将置于此工作表之后。           |
| *Count*  | 可选          | **Variant**  | 要添加的工作表数。默认值为 1。                               |
| *Type*   | 可选          | **Variant**  | 指定工作表类型。可以为下列 XlSheetType 常量之一：xlWorksheet、xlChart、xlExcel4MacroSheet 或 xlExcel4IntlMacroSheet。如果基于现有模板插入工作表，则指定该模板的路径。默认值为 xlWorksheet。 |

**说明**

如果同时省略 *Before* 和 *After*，则新工作表插入到活动工作表之前。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将新建工作表插入到活动工作簿的最后一张工作表之前。*/ function test() {     Application.ActiveWorkbook.Sheets.Add(Application.Worksheets.Item(Worksheets.Count)) }` |

#### **Sheets.Copy**

将工作表复制到工作簿的另一位置。

**语法**

**express.Copy(Before, After)**

*express*   一个代表 **Sheets** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before* | 可选          | **Variant**  | 将要在其之前放置所复制工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选          | **Variant**  | 将要在其之后放置所复制工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After*，则 ET 将新建一个工作簿，其中包含复制的工作表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例复制工作表 Sheet1，并将其放置在工作表 Sheet3 之前。*/ function test() {     Application.Worksheets.Item("Sheet1").Copy(Application.Worksheets.Item("Sheet3")) }` |

#### **Sheets.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Sheets** 对象的变量。

#### **Sheets.FillAcrossSheets**

将单元格区域复制到集合中所有其他工作表的同一位置。

**语法**

**express.FillAcrossSheets(Range, Type)**

*express*   一个代表 **Sheets** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**   | **说明**                                                     |
| -------- | ------------- | -------------- | ------------------------------------------------------------ |
| *Range*  | 必选          | **Range**      | 要填充到集合中所有工作表上的单元格区域。该区域必须来自集合中的某个工作表。 |
| *Type*   | 可选          | **XlFillWith** | 指定如何复制区域。                                           |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例用工作表 Sheet1 上 A1:C5 单元格区域的内容填充工作表 Sheet5 和 Sheet7 上的相同区域。*/   x = ["Sheet1", "Sheet5", "Sheet7"]   Application.Sheets.Item(x).FillAcrossSheets(Application.Worksheets.Item("Sheet1").Range("A1:C5")) }` |

#### **Sheets.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Sheets** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例激活工作表 Sheet1。*/ Application.Sheets.Item("sheet1").Activate()` |

#### **Sheets.Move**

将工作表移到工作簿中的其他位置。

**语法**

**express.Move(Before, After)**

*express*   一个代表 **Sheets** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Before* | 可选          | **Variant**  | 在其之前放置移动工作表的工作表。如果指定了 After，则不能指定 Before。 |
| *After*  | 可选          | **Variant**  | 在其之后放置移动工作表的工作表。如果指定了 Before，则不能指定 After。 |

**说明**

如果既不指定 *Before* 也不指定 *After*，ET 将新建一个工作簿，其中包含所移动的工作表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将当前活动工作簿的 Sheet1 移到 Sheet3 之后。*/ Application.Worksheets.Item("Sheet1").Move(null,Application.Worksheets.Item("Sheet3"))` |

#### **Sheets.PrintOut**

打印对象。返回Variant值

**语法**

**express.PrintOut(From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName, IgnorePrintAreas)**

*express*   一个代表 **Sheets** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *From*             | 可选          | **Variant**  | 打印的开始页号。如果省略此参数，则从起始位置开始打印。       |
| *To*               | 可选          | **Variant**  | 打印的终止页号。如果省略此参数，则打印至最后一页。           |
| *Copies*           | 可选          | **Variant**  | 打印份数。如果省略此参数，则只打印一份。                     |
| *Preview*          | 可选          | **Variant**  | 如果为 True，ET 将在打印对象之前调用打印预览。如果为 False（或省略该参数），则立即打印对象。 |
| *ActivePrinter*    | 可选          | **Variant**  | 设置活动打印机的名称。                                       |
| *PrintToFile*      | 可选          | **Variant**  | 如果为 True，则打印到文件。如果没有指定 PrToFileName，ET 将提示用户输入要使用的输出文件的文件名。 |
| *Collate*          | 可选          | **Variant**  | 如果为 True，则逐份打印多个副本。                            |
| *PrToFileName*     | 可选          | **Variant**  | 如果 PrintToFile 设为 True，则该参数指定要打印到的文件名。   |
| *IgnorePrintAreas* | 可选          | **Variant**  | 如果为 True，则忽略打印区域并打印整个对象。                  |

**说明**

*From* 和 *To* 所描述的“页”指的是要打印的页，并非指定工作表或工作簿中的全部页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例打印当前活动工作表。*/ Application.ActiveSheet.PrintOut()` |

#### **Sheets.PrintPreview**

按对象打印后的外观效果显示对象的预览。

**语法**

**express.PrintPreview(EnableChanges)**

*express*   一个代表 **Sheets** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *EnableChanges* | 可选          | **Variant**  | 传递 Boolean 值，以指定用户是否可更改边距和打印预览中可用的其他页面设置选项。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在打印预览中显示工作表 Sheet1。*/ Application.Worksheets.Item("Sheet1").PrintPreview()` |

#### **Sheets.Select**

选择对象。

**语法**

**express.Select(Replace)**

*express*   一个代表 **Sheets** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Replace* | 可选          | **Variant**  | （仅用于工作表）。如果为 True，则用指定的对象替换当前所选内容。如果为 False，则扩展当前所选内容以包括以前选择的对象和指定的对象。 |

**成员属性**

#### **Sheets.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Sheets** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET"){         alert("This is an ET Application object.")     }     else{         alert("This is not an ET Application object.")     } }` |

#### **Sheets.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Sheets** 对象的变量。

#### **Sheets.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Sheets** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Sheets.HPageBreaks**

返回一个 **HPageBreaks**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/HPageBreaks/HPageBreaks%20.htm#jsObject_HPageBreaks)集合，它代表工作表上的水平分页符。只读。

**语法**

**express.HPageBreaks**

*express*   一个代表 **Sheets** 对象的变量。

**说明**

每张工作表最多有 1026 个水平分页符。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例显示全屏水平分页符和打印区水平分页符的个数。*/   let cFull = 0   let cPartial = 0   for(let i = 1; i <= Application.Worksheets.Item(1).HPageBreaks.Count; i++){       if(Application.Worksheets.Item(1).HPageBreaks.Item(i).Extent == xlPageBreakFull){           cFull = cFull + 1       }       else{           cPartial = cPartial + 1       }   }   alert(cFull + " full-screen page breaks, " + cPartial + " print-area page breaks") }` |

#### **Sheets.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Sheets** 对象的变量。

#### **Sheets.VPageBreaks**

返回一个 **VPageBreaks**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/VPageBreaks/VPageBreaks%20.htm#jsObject_VPageBreaks)集合，它代表工作表上的垂直分页符。只读。

**语法**

**express.VPageBreaks**

*express*   一个代表 **Sheets** 对象的变量。

#### **Sheets.Visible**

返回或设置一个 **Variant** 值，它确定对象是否可见。

**语法**

**express.Visible**

*express*   一个代表 **Sheets** 对象的变量。

适用环境：web

适用平台：windows/linux