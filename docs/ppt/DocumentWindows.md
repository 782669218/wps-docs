**DocumentWindows**



WPP中当前所有打开的 [DocumentWindow](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/DocumentWindow/DocumentWindow%20.htm#jsObject_DocumentWindow) 对象的集合。该集合不包含打开的幻灯片放映窗口，这些窗口包含在 [SlideShowWindows](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/SlideShowWindows/SlideShowWindows%20.htm#jsObject_SlideShowWindows) 集合中。

**方法**

|                                                              | 名称        | 说明                     |
| ------------------------------------------------------------ | ----------- | ------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Arrange** | 指定是层叠还是平铺窗口   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**    | 从指定集合中返回单个对象 |

**属性**

|                                                              | 名称            | 说明                                                |
| ------------------------------------------------------------ | --------------- | --------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个Application对象，该对象表示指定对象的创建者 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象                                |

**成员方法**

#### **DocumentWindows.Arrange**

指定是层叠还是平铺窗口

**语法**

**express.Arrange(arrangeStyle)**

*express*   一个代表 **DocumentWindows** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**       | **说明**               |
| -------------- | ------------- | ------------------ | ---------------------- |
| *arrangeStyle* | 可选          | **PpArrangeStyle** | 指定是层叠还是平铺窗口 |

#### **DocumentWindows.Item**

从指定集合中返回单个对象

**语法**

**express.Item(Index)**

*express*   一个代表 **DocumentWindows** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                       |
| -------- | ------------- | ------------ | ------------------------------ |
| *Index*  | 必选          | **Long**     | 集合中要返回的单个对象的索引号 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭第二个文档窗口*/ Application.Windows.Item(2).Close()` |

**成员属性**

#### **DocumentWindows.Application**

返回一个Application对象，该对象表示指定对象的创建者

**语法**

**express.Application**

*express*   一个代表 **DocumentWindows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个Presentation对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行的WPP的文件夹中。*/ function test(){      let pptPres = Application.ActivePresentation      pptPres.Slides.Add(1,1)      pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

#### **DocumentWindows.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **DocumentWindows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口*/ function test(){     let dWindows = Application.Windows      for（let i = 2; i<= dWindows.Count; i++){           dWindows.Item(i).Close()    } }` |

#### **DocumentWindows.Parent**

返回指定对象的父对象

**语法**

**express.Parent**

*express*   一个代表 **DocumentWindows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。 function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux