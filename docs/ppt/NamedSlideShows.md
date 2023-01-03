**NamedSlideShows**



演示文稿中所有 **NamedSlideShow** 对象的集合。每个 **NamedSlideShow** 对象代表一个自定义幻灯片放映。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 创建一个新的命名幻灯片放映，并将其添加到指定演示文稿的命名幻灯片放映集合中。返回一个代表新命名幻灯片放映的 **NamedSlideShow** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从指定集合中返回单个对象。                                   |

**属性**

|                                                              | 名称            | 说明                                                        |
| ------------------------------------------------------------ | --------------- | ----------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                      |

**成员方法**

#### **NamedSlideShows.Add**

创建一个新的命名幻灯片放映，并将其添加到指定演示文稿的命名幻灯片放映集合中。返回一个代表新命名幻灯片放映的 **NamedSlideShow** 对象。

**语法**

**express.Add(Name, safeArrayOfSlideIDs)**

*express*   一个代表 **NamedSlideShows** 对象的变量。

**参数**

| **名称**              | **必选/可选** | **数据类型** | **说明**                                          |
| --------------------- | ------------- | ------------ | ------------------------------------------------- |
| *Name*                | 必选          | **String**   | 幻灯片放映的名称。                                |
| *safeArrayOfSlideIDs* | 必选          | **Variant**  | 包含将在幻灯片放映中显示的幻灯片的唯一幻灯片 ID。 |

**返回值**

NamedSlideShow

**说明**

添加命名幻灯片放映时指定的名称就是用作 **Run** 方法的参数的名称，以便运行命名幻灯片放映。

#### **NamedSlideShows.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **NamedSlideShows** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

NamedSlideShow

**成员属性**

#### **NamedSlideShows.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **NamedSlideShows** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **NamedSlideShows.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **NamedSlideShows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口。*/ function test(){ 　　　　let win = Application.Windows    　　　　 for(let i = 2; i <= win.Count; i++){      　　　　   win.Item(i).Close()   　　　　  } }` |

#### **NamedSlideShows.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **NamedSlideShows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes 　　　　let sha = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　sha.TextRange.Text = "Test text" 　　　　sha.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux