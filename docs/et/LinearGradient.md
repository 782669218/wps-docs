**LinearGradient**



**LinearGradient** 对象沿特定角度以线性方式在一系列颜色间转换。

**说明**

| ![img]()注释                                       |
| -------------------------------------------------- |
| 当使用 **LinearGradient** 对象时，应考虑以下几点： |

- 试图访问不具有现有渐变填充的 **Interior** 对象的 Gradient 属性会引起运行时错误。访问 Gradient 属性之前请注意 `Interior.Pattern` 属性。
- 如果将 Interior.Pattern 从渐变类型更改为非渐变类型，Gradient 对象将采用默认值。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorStops**  | 返回 **LinearGradient** 对象的 **ColorStops**。只读。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Degree**      | 选定区域中线性渐变填充的角度。可读/写。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员属性**

#### **LinearGradient.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **LinearGradient** 对象的变量。

#### **LinearGradient.ColorStops**

返回 **LinearGradient** 对象的 **ColorStops**。只读。

**语法**

**express.ColorStops**

*express*   一个代表 **LinearGradient** 对象的变量。

#### **LinearGradient.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **LinearGradient** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。Creator 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **LinearGradient.Degree**

选定区域中线性渐变填充的角度。可读/写。

**语法**

**express.Degree**

*express*   一个代表 **LinearGradient** 对象的变量。

**说明**

使用 0 - 360 之间的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ Selection.Interior.Pattern = xlPatternLinearGradient Selection.Interior.Gradient.Degree = 45 }` |

#### **LinearGradient.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **LinearGradient** 对象的变量。

适用环境：web

适用平台：windows/linux