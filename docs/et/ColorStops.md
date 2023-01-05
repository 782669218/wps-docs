**ColorStops**



指定的数据系列中所有 ColorStop 对象的集合。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示具有 LinearGradients 的 ColorStops。*/ function test(){ let myInterior = Selection.Interior     myInterior.Pattern = xlPatternLinearGradient     myInterior.Gradient.Degree = 90     myInterior.Gradient.ColorStops.Clear()      //adds stops after any have been cleared let myColorStops1 = Selection.Interior.Gradient.ColorStops.Add(0)     myColorStops1.ThemeColor = xlThemeColorDark1     myColorStops1.TintAndShade = 0  let myColorStops2 = Selection.Interior.Gradient.ColorStops.Add(1)     myColorStops2.ThemeColor = xlThemeColorAccent1     myColorStops2.TintAndShade = 0 }` |

每个 **ColorStop** 对象代表一个区域或选定内容中渐变填充的一个颜色光圈。

**方法**

|                                                              | 名称        | 说明 |
| ------------------------------------------------------------ | ----------- | ---- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reserve** |      |

**属性**

|                                                              | 名称        | 说明 |
| ------------------------------------------------------------ | ----------- | ---- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reserve** |      |

**成员方法**

#### **ColorStops.Reserve**

**语法**

**express.Reserve()**

*express*   一个代表 **ColorStops** 对象的变量。

**成员属性**

#### **ColorStops.Reserve**

**语法**

**express.Reserve**

*express*   一个代表 **ColorStops** 对象的变量。

适用环境：web

适用平台：windows/linux