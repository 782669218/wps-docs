**FormatColor**



代表为色阶条件格式阈值指定的填充色或数据条条件格式的条形颜色。

**说明**

您可以通过传递 [**Color** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatColor/FormatColor%20.htm#FormatColor.Color)属性中的 RGB 值来选择颜色，或者通过使用 [**ThemeColor** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/FormatColor/FormatColor%20.htm#FormatColor.ThemeColor)属性在主题调色板中编制索引来指定颜色。

以下代码示例创建了一个数字范围，然后将双色色阶条件格式规则应用于该范围。然后通过在 **ColorScaleCriteria** 集合中编制索引来设置单独的条件，从而指定最小阈值的颜色为红色，最大阈值的颜色为蓝色。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     //Fill cells with sample data from 1 to 10     let sheet = ActiveSheet     sheet.Range("C1").Value2 = 1     sheet.Range("C2").Value2 = 2     sheet.Range("C1:C2").AutoFill(Range("C1:C10"))          Range("C1:C10").Select()          //Create a two-color ColorScale object for the created sample data range     let cfColorScale = Selection.FormatConditions.AddColorScale(2)          //Set the minimum threshold to red and maximum threshold to blue     cfColorScale.ColorScaleCriteria.Item(1).FormatColor.Color = (255, 0, 0)     cfColorScale.ColorScaleCriteria.Item(2).FormatColor.Color = (0, 0, 255) }` |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Color**        | 返回或设置与数据条或色阶条件格式规则的阈值关联的填充色。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorIndex**   | 返回或设置 **XlColorIndex**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlColorIndex%20%E6%9E%9A%E4%B8%BE.html)枚举的常量之一，指定是否以当前调色板中颜色的索引值形式表示填充色。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThemeColor**   | 返回或设置一个 **XlThemeColor**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlThemeColor%20%E6%9E%9A%E4%B8%BE.html)枚举常量，该值指定数据条或色阶条件格式阈值中所用的主题颜色。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TintAndShade** | 返回或设置一个 **Single** 类型的值，该值使用于数据条或色阶条件格式规则的单元格的填充颜色变浅或变深。 |

**成员属性**

#### **FormatColor.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **FormatColor** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **FormatColor.Color**

返回或设置与数据条或色阶条件格式规则的阈值关联的填充色。

**语法**

**express.Color**

*express*   一个代表 **FormatColor** 对象的变量。

**说明**

格式颜色以 RGB 函数表示。例如，要将颜色设置为红色，请使用 `RGB(255,0,0)`。

#### **FormatColor.ColorIndex**

返回或设置 **XlColorIndex**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlColorIndex%20%E6%9E%9A%E4%B8%BE.html)枚举的常量之一，指定是否以当前调色板中颜色的索引值形式表示填充色。

**语法**

**express.ColorIndex**

*express*   一个代表 **FormatColor** 对象的变量。

**说明**

此属性用于色阶或数据条条件格式规则中的每个阈值。

#### **FormatColor.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **FormatColor** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **FormatColor.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **FormatColor** 对象的变量。

#### **FormatColor.ThemeColor**

返回或设置一个 **XlThemeColor**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlThemeColor%20%E6%9E%9A%E4%B8%BE.html)枚举常量，该值指定数据条或色阶条件格式阈值中所用的主题颜色。

**语法**

**express.ThemeColor**

*express*   一个代表 **FormatColor** 对象的变量。

#### **FormatColor.TintAndShade**

返回或设置一个 **Single** 类型的值，该值使用于数据条或色阶条件格式规则的单元格的填充颜色变浅或变深。

**语法**

**express.TintAndShade**

*express*   一个代表 **FormatColor** 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

适用环境：web

适用平台：windows/linux