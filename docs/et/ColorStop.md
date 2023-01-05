**ColorStop**



表示某一区域或选定内容中渐变填充的光圈点。

**说明**

**ColorStop** 对象可用来设置单元格填充属性，包括 **Color**、**ThemeColor** 和 **TintAndShade** 属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例演示如何对 ColorStop 对象应用属性。*/ function test(){ let myInterior = Application.Selection.Interior     myInterior.Pattern = xlPatternLinearGradient     myInterior.Gradient.Degree = 135     myInterior.Gradient.ColorStops.Clear()  let myColorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)     myColorStops1.ThemeColor = xlThemeColorDark1     myColorStops1.TintAndShade = 0  let myColorStops2 = Application.Selection.Interior.Gradient.ColorStops.Add(0.5)     myColorStops2.ThemeColor = xlThemeColorAccent1     myColorStops2.TintAndShade = 0  let myColorStops3 = Application.Selection.Interior.Gradient.ColorStops.Add(1)     myColorStops3.ThemeColor = xlThemeColorDark1     myColorStops3.TintAndShade = 0 }` |

**方法**

|                                                              | 名称       | 说明           |
| ------------------------------------------------------------ | ---------- | -------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定对象。 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Color**        | 返回或设置指定对象的颜色。可读/写。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**     | 返回或设置 **ColorStop** 对象的位置。可读/写。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThemeColor**   | 返回或设置指定对象的主颜色。可读/写。                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TintAndShade** | 返回或设置指定对象的淡色和阴影。可读/写。                    |

**成员方法**

#### **ColorStop.Delete**

删除指定对象。

**语法**

**express.Delete()**

*express*   一个代表 **ColorStop** 对象的变量。

**成员属性**

#### **ColorStop.Application**

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **ColorStop** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ let myObject = Application.ActiveWorkbook if(myObject.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") } }` |

#### **ColorStop.Color**

返回或设置指定对象的颜色。可读/写。

**语法**

**express.Color**

*express*   一个代表 **ColorStop** 对象的变量。

#### **ColorStop.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ColorStop** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。Creator 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ColorStop.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ColorStop** 对象的变量。

#### **ColorStop.Position**

返回或设置 **ColorStop** 对象的位置。可读/写。

**语法**

**express.Position**

*express*   一个代表 **ColorStop** 对象的变量。

#### **ColorStop.ThemeColor**

返回或设置指定对象的主颜色。可读/写。

**语法**

**express.ThemeColor**

*express*   一个代表 **ColorStop** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*对当前选定内容应用主题颜色。*/  function test(){ Range("A1:A10").Select() let myColorStop = Application.Selection.Interior.Gradient.ColorStops.Add(1)     myColorStop.ThemeColor = xlThemeColorAccent1     myColorStop.TintAndShade = 0 }` |

#### **ColorStop.TintAndShade**

返回或设置指定对象的淡色和阴影。可读/写。

**语法**

**express.TintAndShade**

*express*   一个代表 **ColorStop** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*对当前选定内容应用淡色和阴影。*/ function test(){ Range("A1:A10").Select() let myColorStop = Application.Selection.Interior.Gradient.ColorStops.Add(1)     myColorStop.ThemeColor = xlThemeColorAccent1     myColorStop.TintAndShade = 0 }` |

适用环境：web

适用平台：windows/linux