**Panes**



**Pane** 对象的集合，这些对象代表普通视图内文档窗口中的幻灯片、大纲和备注窗格，或代表文档窗口中其他任意视图内的单个窗格。

**说明**

在普通视图中，**Panes** 集合包含三个成员。所有其他文档窗口视图则仅有一个窗格，因而 **Panes** 集合也就只有一个成员。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `使用 Panes 属性可返回 Panes 集合。以下示例将测试活动窗口中窗格的数目。如果值为 1，则说明是除普通视图之外的其他视图，接着将激活普通视图，然后使用纵向窗格分隔线将文档窗口进行划分，使大纲窗格占 15% 位置，幻灯片窗格占 85% 位置。 function test() {     let mypane = Application.ActiveWindow      if(mypane.Panes.Count == 1){        mypane.ViewType = ppViewNormal        mypane.SplitHorizontal = 15     } }` |

**方法**

|                                                              | 名称     | 说明                                   |
| ------------------------------------------------------------ | -------- | -------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从指定集合中返回单个对象。返回值为Pane |

**属性**

|                                                              | 名称            | 说明                                                        |
| ------------------------------------------------------------ | --------------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                      |

**成员方法**

#### **Panes.Item**

从指定集合中返回单个对象。返回值为Pane

**语法**

**express.Item(Index)**

*express*   一个代表 **Panes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                         |
| -------- | ------------- | ------------ | -------------------------------- |
| *Index*  | 必选          | **Long**     | 集合中要返回的单个对象的索引号。 |

**成员属性**

#### **Panes.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Panes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {   let mypane = Application.ActivePresentation.Slides.Item(1).Shapes    for(let i = 1; i <= mypane.Count; i++){       if(mypane.Item(i).Type == msoLinkedOLEObject){           alert(mypane.Item(i).OLEFormat.Application.Name)       }   } }` |

#### **Panes.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **Panes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口。*/ function test() {   let mypane = Application.Windows     for(let i = 2; i <= mypane.Count; i++){       mypane.Item(i).Close()   } }` |

#### **Panes.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Panes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。 function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux