**Styles**



指定的数据透视表中所有 **Style**

**说明**

每一个 **Style** 对象都代表对某区域的样式描述。**Style** 对象包含样式的所有属性（字体、数字格式、对齐方式，等等）。有几种内置的样式，包括“常规”、“货币”和“百分比”。

使用 **Styles** 属性可返回 **Styles** 集合。下例创建活动工作簿中工作表一上的样式名的列表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     for(let i = 1;i <= Application.ActiveWorkbook.Styles.Count; i++){         Application.Worksheets.Item(1).Cells.Item(i, 1).Value2 = ActiveWorkbook.Styles.Item(i).Name     } }` |

使用 [**Add** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Styles/Styles%20.htm#Styles.Add)方法可创建一个新的样式并将它添加到集合。下例基于“常规”样式创建一个新的样式，修改边框和字体，然后将该新样式应用到单元格 A25:A30。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){     let rng = Application.ActiveWorkbook.Styles.Add("Bookman Top Border")     rng.Borders.Item(xlTop).LineStyle = xlDouble     rng.Font.Bold = true     rng.Font.Name = "Bookman"     Application.Worksheets.Item(1).Range("A25:A30").Style = "Bookman Top Border" }` |

使用 **Styles**(*index*)（其中 *index* 是样式索引号或名称）可从工作簿的 **Style** 集合中返回一个 **Styles** 对象。下例通过设置活动工作簿中“常规”样式的 **Bold** 属性来更改该样式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true` |

**方法**

|                                                              | 名称      | 说明                                                         |
| ------------------------------------------------------------ | --------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**   | 新建样式并将其添加到当前工作簿的可用样式列表中。返回一个代表新样式的 **Style**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#jsObject_Style)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**  | 从集合中返回一个对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Merge** | 将另一张工作簿中的样式合并到 **Styles** 集合中。返回Variant值 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **Styles.Add**

新建样式并将其添加到当前工作簿的可用样式列表中。返回一个代表新样式的 **Style**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#jsObject_Style)对象。

**语法**

**express.Add(Name)**

*express*   一个代表 **Styles** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Name*   | 必选          | **String**   | 新样式的名称。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例基于工作表 Sheet1 中的单元格 A1 定义新样式。*/ function test(){   let Sty = Application.ActiveWorkbook.Styles.Add("theNewStyle")   Sty.IncludeNumber = false   Sty.IncludeFont = true   Sty.IncludeAlignment = false   Sty.IncludeBorder = false   Sty.IncludePatterns = false   Sty.IncludeProtection = false   Sty.Font.Name = "Arial"   Sty.Font.Size = 18 }` |

#### **Styles.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Styles** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例通过设置“常规”样式的 Bold 属性来更改活动工作簿中的该样式。*/ Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true` |

#### **Styles.Merge**

将另一张工作簿中的样式合并到 **Styles** 集合中。返回Variant值

**语法**

**express.Merge(Workbook)**

*express*   一个代表 **Styles** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                           |
| ---------- | ------------- | ------------ | -------------------------------------------------- |
| *Workbook* | 必选          | **Variant**  | 一个 Workbook 对象，它代表包含待合并样式的工作簿。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将工作簿 Template.xls 中的样式合并到活动工作簿中。*/ Application.ActiveWorkbook.Styles.Merge(Workbooks.Item("TEMPLATE.XLS"))` |

**成员属性**

#### **Styles.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Styles** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例显示一条有关创建 myObject 的应用程序的消息。*/   let myObject = Application.ActiveWorkbook   if(myObject.Application.Value == "ET"){       alert("This is an ET Application object.")   }   else{       alert("This is not an ET Application object.")   } }` |

#### **Styles.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Styles** 对象的变量。

#### **Styles.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Styles** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Styles.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Styles** 对象的变量。

适用环境：web

适用平台：windows/linux