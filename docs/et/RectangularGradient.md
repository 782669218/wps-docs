**RectangularGradient**



**RectangularGradient** 对象沿特定角度以线性方式在一系列颜色间转换。

**说明**

- 试图访问不具有现有渐变填充的 **Interior** 对象的 Gradient 属性会引起运行时错误。访问 Gradient 属性之前请注意 `Interior.Pattern` 属性。
- 如果将 Interior.Pattern 从渐变类型更改为非渐变类型，Gradient 对象将采用默认值进行填充。

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorStops**      | 返回 **RectangularGradient** 对象的 **ColorStops** 集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RectangleBottom** | 代表渐变填充收敛到的点或矢量。可读/写。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RectangleLeft**   | 代表渐变填充收敛到的点或矢量。可读/写。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RectangleRight**  | 代表渐变填充收敛到的点或矢量。可读/写。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RectangleTop**    | 代表渐变填充收敛到的点或矢量。可读/写。                      |

**成员属性**

#### **RectangularGradient.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **RectangularGradient** 对象的变量。

#### **RectangularGradient.ColorStops**

返回 **RectangularGradient** 对象的 **ColorStops** 集合。只读。

**语法**

**express.ColorStops**

*express*   一个代表 **RectangularGradient** 对象的变量。

#### **RectangularGradient.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **RectangularGradient** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。Creator 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **RectangularGradient.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **RectangularGradient** 对象的变量。

#### **RectangularGradient.RectangleBottom**

代表渐变填充收敛到的点或矢量。可读/写。

**语法**

**express.RectangleBottom**

*express*   一个代表 **RectangularGradient** 对象的变量。

**说明**

与 RectangleLeft、RectangleRight 和 RectangleTop 一起使用。下表中列出了有效值。

| 属性            | 值   |
| --------------- | ---- |
| RectangleLeft   | 0-1  |
| RectangleRight  | 0-1  |
| RectangleTop    | 0-1  |
| RectangleBottom | 0-1  |

#### **RectangularGradient.RectangleLeft**

代表渐变填充收敛到的点或矢量。可读/写。

**语法**

**express.RectangleLeft**

*express*   一个代表 **RectangularGradient** 对象的变量。

**说明**

与 RectangleRight、RectangleTop 和 RectangleBottom 一起使用。下表中列出了有效值。

| 属性            | 值   |
| --------------- | ---- |
| RectangleLeft   | 0-1  |
| RectangleRight  | 0-1  |
| RectangleTop    | 0-1  |
| RectangleBottom | 0-1  |

#### **RectangularGradient.RectangleRight**

代表渐变填充收敛到的点或矢量。可读/写。

**语法**

**express.RectangleRight**

*express*   一个代表 **RectangularGradient** 对象的变量。

**说明**

与 RectangleLeft、RectangleTop 和 RectangleBottom 一起使用。下表中列出了有效值。

| 属性            | 值   |
| --------------- | ---- |
| RectangleLeft   | 0-1  |
| RectangleRight  | 0-1  |
| RectangleTop    | 0-1  |
| RectangleBottom | 0-1  |

#### **RectangularGradient.RectangleTop**

代表渐变填充收敛到的点或矢量。可读/写。

**语法**

**express.RectangleTop**

*express*   一个代表 **RectangularGradient** 对象的变量。

**说明**

与 RectangleLeft、RectangleRight 和 RectangleBottom 一起使用。下表中列出了有效值。

| 属性            | 值   |
| --------------- | ---- |
| RectangleLeft   | 0-1  |
| RectangleRight  | 0-1  |
| RectangleTop    | 0-1  |
| RectangleBottom | 0-1  |

适用环境：web

适用平台：windows/linux