**ObjectVerbs**



代表指定 OLE 对象的 OLE 动作的集合。OLE 动作是 OLE 对象支持的操作。常用的 OLE 动作有“play”和“edit”。

**方法**

|                                                              | 名称     | 说明                       |
| ------------------------------------------------------------ | -------- | -------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从指定集合中返回单个对象。 |

**属性**

|                                                              | 名称            | 说明                                                        |
| ------------------------------------------------------------ | --------------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                      |

**成员方法**

#### **ObjectVerbs.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **ObjectVerbs** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                         |
| -------- | ------------- | ------------ | -------------------------------- |
| *Index*  | 必选          | **Long**     | 集合中要返回的单个对象的索引号。 |

**返回值**

String

**成员属性**

#### **ObjectVerbs.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **ObjectVerbs** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ 　　　　for(let i = 1; i <= ActivePresentation.Slides.Item(1).Shapes.Count; i++){ 　　　　    if(ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject){ 　　　　        MsgBox(shpOle.OLEFormat.Application.Name) 　　　　    } 　　　　} }` |

#### **ObjectVerbs.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **ObjectVerbs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口。*/ function test(){ 　　　　let win = Application.Windows 　　　　    for(let i = 2; i <= win.Count; i++){ 　　　　        win.Item(2).Close()   　　　　  } }` |

#### **ObjectVerbs.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ObjectVerbs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = ActivePresentation.Slides.Item(1).Shapes 　　　　let sha = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　sha.TextRange.Text = "Test text" 　　　　sha.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux