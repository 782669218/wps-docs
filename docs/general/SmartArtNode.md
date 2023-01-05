#### **SmartArtNode**



Smart Art 图形的数据模型内部的单个语义节点。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码返回 Smart Art 图表中的节点数目。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.AllNodes.Count` |

**方法**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddNode**     | 通过SmartArtNodeType类型的SmartArtNodePosition值指定的方式，将新的SmartArtNode添加到数据模型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**      | 删除当前的 SmartArt 节点。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Demote**      | 将当前节点在数据模型中降低一级。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Larger**      | 增加 SmartArt 节点的大小。模拟 SmartArt 的 WPS Office Fluent 功能区“格式”选项卡上的“增大”按钮的行为。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Promote**     | 将当前节点（以及它的所有子级）在数据模型中提升一级。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReorderDown** | 与项目符号列表中的下一个节点交换节点。此方法会对整个节点系列进行重新排序。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReorderUp**   | 与项目符号列表中的上一个节点交换节点。此方法会对整个节点系列进行重新排序。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Smaller**     | 减小 SmartArt 的大小。模拟 SmartArt 的 WPS Office Fluent 功能区用户界面的“格式”选项卡上的“减小”按钮的行为。 |

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**    | 获取一个 **Application** 对象，该对象代表 **SmartArtNode** 对象的容器应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**        | 获取一个 32 位整数，该整数指示在其中创建了 **SmartArtNode** 对象的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hidden**         | 如果此节点在数据模型中是一个隐藏节点，则返回 **true**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Level**          | 检索节点在层次结构中的级别。只读。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Nodes**          | 检索与此 Smart Art 节点关联的子节点。只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OrgChartLayout** | 如果有一个节点，则检索或设置与此节点关联的 **MsoOrgChartLayoutType**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**         | 返回调用对象。只读。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentNode**     | 检索此 SmartArtNode 的父 SmartArtNode。只读。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shapes**         | 返回与此 **SmartArtNode** 对象关联的形状范围。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFrame2**     | 返回与该 **SmartArtNode** 对象关联的文本。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**           | 检索 SmartArt 节点的类型。只读。                             |

**成员方法**

#### **SmartArtNode.AddNode**

通过SmartArtNodeType类型的SmartArtNodePosition值指定的方式，将新的SmartArtNode添加到数据模型。

**语法**

**express.AddNode(Position, Type)**

*express*   一个代表 **SmartArtNode** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型**                | **说明**                                                     |
| ---------- | ------------- | --------------------------- | ------------------------------------------------------------ |
| *Position* | 可选          | **MsoSmartArtNodePosition** | 指定 SmartArtNode 在数据模型中的位置。例如，msoSmartArtNodeAbove 或 msoSmartArtNodeAfter。 |
| *Type*     | 可选          | **MsoSmartArtNodeType**     | 指定添加的 SmartArtNode 的类型。例如，msoSmartArtNodeTypeAssistant 或 msoSmartArtNodeTypeDefault。 |

**返回值**

SmartArtNode

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码将默认的 SmartArtNode 添加到当前节点下。*/ function test(){ 	var saNode = Application.ActiveWorkbook.ActiveChart.Shapes.Item(1).SmartArt.AllNodes.Item(1); 	saNode = saNode.AddNode(Application.Enum.msoSmartArtNodeBelow, Application.Enum.msoSmartArtNodeTypeDefault); }` |

#### **SmartArtNode.Delete**

删除当前的 SmartArt 节点。

**语法**

**express.Delete()**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

删除节点时，第一个子级得到提升。在下面的数据模型中：

- A

- - B

  - - C

- D

如果删除了 B，则该数据模型的外观如下所示：

- A

- - C

- D

#### **SmartArtNode.Demote**

将当前节点在数据模型中降低一级。

**语法**

**express.Demote()**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

在内容窗格中工作时，此功能模拟 WPS Office Fluent 功能区用户界面中的“降级”按钮。例如，以下面的数据模型为例：

- A

- B

- - C

- D

如果使 B 降级，产生的数据模型的外观如下所示：

- A

- - B
  - C

- D

#### **SmartArtNode.Larger**

增加 SmartArt 节点的大小。模拟 SmartArt 的 WPS Office Fluent 功能区“格式”选项卡上的“增大”按钮的行为。

**语法**

**express.Larger()**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

实际的增加量取决于节点的布局。

#### **SmartArtNode.Promote**

将当前节点（以及它的所有子级）在数据模型中提升一级。

**语法**

**express.Promote()**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

在内容窗格中工作时，此功能模拟 WPS Office Fluent 功能区用户界面中的“提升”按钮。例如，以下面的数据模型为例：

- A

- - B

  - - C

- D

如果使 B 提升，产生的数据模型的外观如下所示：

- A

- B

- - C

- D

#### **SmartArtNode.ReorderDown**

与项目符号列表中的下一个节点交换节点。此方法会对整个节点系列进行重新排序。

**语法**

**express.ReorderDown()**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

此方法模拟单击 WPS Office Fluent 功能区用户界面上的“向下重新排序”按钮，该按钮位于“向下重新排序”的“设计”组的“SmartArt 工具”选项卡上。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码将第一个节点与下一个节点交换，并对其所有后代进行重新排序。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Item(1).ReorderDown()` |

#### **SmartArtNode.ReorderUp**

与项目符号列表中的上一个节点交换节点。此方法会对整个节点系列进行重新排序。

**语法**

**express.ReorderUp()**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

此方法模拟单击 WPS Office Fluent 功能区用户界面上的“向上重新排序”按钮，该按钮位于“向上重新排序”的“设计”组的“SmartArt 工具”选项卡上。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码将第一个节点与下一个节点交换，并对父节点进行重新排序。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Item(1).ReorderUp()` |

#### **SmartArtNode.Smaller**

减小 SmartArt 的大小。模拟 SmartArt 的 WPS Office Fluent 功能区用户界面的“格式”选项卡上的“减小”按钮的行为。

**语法**

**express.Smaller()**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

减小量取决于节点的布局大小。

**成员属性**

#### **SmartArtNode.Application**

获取一个 **Application** 对象，该对象代表 **SmartArtNode** 对象的容器应用程序。只读。

**语法**

**express.Application**

*express*   一个代表 **SmartArtNode** 对象的变量。

#### **SmartArtNode.Creator**

获取一个 32 位整数，该整数指示在其中创建了 **SmartArtNode** 对象的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **SmartArtNode** 对象的变量。

#### **SmartArtNode.Hidden**

如果此节点在数据模型中是一个隐藏节点，则返回 **true**。只读。

**语法**

**express.Hidden**

*express*   一个代表 **SmartArtNode** 对象的变量。

#### **SmartArtNode.Level**

检索节点在层次结构中的级别。只读。

**语法**

**express.Level**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

起始级别为 1 级并且向上递增。如果某个节点没有级别，则返回 0。例如，在下面的数据模型中，A 和 F 为 1 级，B 和 D 为 2 级，C 和 E 为 3 级。

- A

- - B

  - - C

  - D

- - E

- F

#### **SmartArtNode.Nodes**

检索与此 Smart Art 节点关联的子节点。只读。

**语法**

**express.Nodes**

*express*   一个代表 **SmartArtNode** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码返回 Smart Art 图表中的节点数目。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Item(1).Nodes.Count` |

#### **SmartArtNode.OrgChartLayout**

如果有一个节点，则检索或设置与此节点关联的 **MsoOrgChartLayoutType**。可读/写。

**语法**

**express.OrgChartLayout**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

以下是可能的成员：

- msoOrgChartLayoutBothHanging
- msoOrgChartLayoutDefault
- msoOrgChartLayoutLeftHanging
- msoOrgChartLayoutMixed
- msoOrgChartLayoutRightHanging
- msoOrgChartLayoutStandard

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码将 OrgChartLayout 属性设置为默认布局。*/ function test(){ 	var saNode =  Application.ActiveWorkbook.ActiveChart.Shapes.Item(1); 	saNode.OrgChartLayout = Application.Enum.msoOrgChartLayoutDefault; }` |

#### **SmartArtNode.Parent**

返回调用对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **SmartArtNode** 对象的变量。

#### **SmartArtNode.ParentNode**

检索此 SmartArtNode 的父 SmartArtNode。只读。

**语法**

**express.ParentNode**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

如果此节点是根节点，则返回一个错误。

#### **SmartArtNode.Shapes**

返回与此 **SmartArtNode** 对象关联的形状范围。只读。

**语法**

**express.Shapes**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

若要查找该范围，请使用 Count 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码返回形状范围。*/ Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt.Nodes.Item(1).Shapes.Count` |

#### **SmartArtNode.TextFrame2**

返回与该 **SmartArtNode** 对象关联的文本。只读。

**语法**

**express.TextFrame2**

*express*   一个代表 **SmartArtNode** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例设置第一个节点内部的文本。*/ function test(){ 	var smartart = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).SmartArt 	smartart.AllNodes.Item(1).TextFrame2.TextRange.Text="Node 1" }` |

#### **SmartArtNode.Type**

检索 SmartArt 节点的类型。只读。

**语法**

**express.Type**

*express*   一个代表 **SmartArtNode** 对象的变量。

**说明**

可能值为 msoSmartArtNodeTypeAssistant 或 msoSmartArtNodeTypeDefault。

适用环境：web

适用平台：windows/linux