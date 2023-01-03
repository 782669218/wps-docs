#### **SmartArt**



用于与 SmartArt 图形交互的顶级类。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码将顶级节点添加到带有项目符号的文本窗格中。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Add()` |

**方法**

|                                                              | 名称      | 说明                               |
| ------------------------------------------------------------ | --------- | ---------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset** | 将 SmartArt 图形重置为其原始状态。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllNodes**    | 检索包含 SmartArt 图表内部所有节点的 **SmartArtNodes** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，该对象代表 **SmartArt** 对象的容器应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Color**       | 检索或设置应用到 Smart Art 图形的 Smart Art 颜色样式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，该整数指示在其中创建了 **SmartArt** 对象的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Layout**      | 检索或设置与 Smart Art 图形关联的 Smart Art 布局。可读/写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Nodes**       | 检索 SmartArt 图表中根节点的子级。只读。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回调用对象。只读。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **QuickStyle**  | 检索或设置应用到 SmartArt 图形的 SmartArt 快速样式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reverse**     | 获取或设置 SmartArt 图表与（从左到右）LTR 或（从右到左）RTL（如果该图表支持反向）有关的状态。可读/写。 |

**成员方法**

#### **SmartArt.Reset**

将 SmartArt 图形重置为其原始状态。

**语法**

**express.Reset()**

*express*   一个代表 **SmartArt** 对象的变量。

**说明**

此操作对应于 SmartArt 的“设计”选项卡下的“重设图形”按钮提供的功能。

**成员属性**

#### **SmartArt.AllNodes**

检索包含 SmartArt 图表内部所有节点的 **SmartArtNodes** 对象。只读。

**语法**

**express.AllNodes**

*express*   一个代表 **SmartArt** 对象的变量。

**说明**

节点是按照顺序检索的，与数据模型无关。例如，下面的数据模型将按照 A、B、C、D、E、F 的顺序检索节点。

- A

- - B

  - - C

  - D

- - - E

- F

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例设置第一个节点内部的文本。*/Application.ActiveWorkbook.ActiveChart.Shapes.Item(1).SmartArt.AllNodes.Item(1).TextFrame2.TextRange.Text="Node 1"` |

#### **SmartArt.Application**

获取一个 **Application** 对象，该对象代表 **SmartArt** 对象的容器应用程序。只读。

**语法**

**express.Application**

*express*   一个代表 **SmartArt** 对象的变量。

#### **SmartArt.Color**

检索或设置应用到 Smart Art 图形的 Smart Art 颜色样式。可读/写。

**语法**

**express.Color**

*express*   一个代表 **SmartArt** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码设置 Smart Art 图表的配色方案。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Color = Application.SmartArtColors.Item(1)` |

#### **SmartArt.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **SmartArt** 对象的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **SmartArt** 对象的变量。

#### **SmartArt.Layout**

检索或设置与 Smart Art 图形关联的 Smart Art 布局。可读/写。

**语法**

**express.Layout**

*express*   一个代表 **SmartArt** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码设置 Smart Art 布局。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Layout = Application.SmartArtLayouts.Item(1)` |

#### **SmartArt.Nodes**

检索 SmartArt 图表中根节点的子级。只读。

**语法**

**express.Nodes**

*express*   一个代表 **SmartArt** 对象的变量。

**说明**

根节点没有父节点并且仅包含子级（如果 SmartArt 图形的数据模型内有子级存在）。在下面的示例中，将返回节点 A 和 F。

- A

- - B

  - - C

  - D

- - - E

- F

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码在 WPP 中添加一个顶级节点。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Add()` |

#### **SmartArt.Parent**

返回调用对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **SmartArt** 对象的变量。

#### **SmartArt.QuickStyle**

检索或设置应用到 SmartArt 图形的 SmartArt 快速样式。可读/写。

**语法**

**express.QuickStyle**

*express*   一个代表 **SmartArt** 对象的变量。

**示例**

------

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码更改 WPP 中 Smart Art 的快速样式。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(1)` |

#### **SmartArt.Reverse**

获取或设置 SmartArt 图表与（从左到右）LTR 或（从右到左）RTL（如果该图表支持反向）有关的状态。可读/写。

**语法**

**express.Reverse**

*express*   一个代表 **SmartArt** 对象的变量。

适用环境：web

适用平台：windows/linux