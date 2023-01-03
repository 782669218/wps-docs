**PageSetup**



包含演示文稿中有关幻灯片、备注页、讲义及大纲的页面设置的信息。

**说明**

使用 **PageSetup** 属性返回 **PageSetup** 对象。以下示例将活动演示文稿中所有幻灯片设为宽 11 英寸、高 8.5 英寸，并将演示文稿的幻灯片编号设为从 17 开始。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let page = Application.ActivePresentation.PageSetup     page.SlideWidth = 11 * 72     page.SlideHeight = 8.5 * 72     page.FirstSlideNumber = 2 }` |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstSlideNumber** | 返回或设置演示文稿中第一张幻灯片的编号。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NotesOrientation** | 返回或设置指定演示文稿的备注页、讲义和大纲的屏幕显示和打印方向。可读/写 **MsoOrientation** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**           | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideHeight**      | 以磅为单位返回或设置幻灯片的高度。可读/写 **Single** 类型。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideOrientation** | 返回或设置指定演示文稿中幻灯片的屏幕显示和打印方向。可读/写 **MsoOrientation** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideSize**        | 返回或设置指定演示文稿的幻灯片大小。可读/写 **PpSlideSizeType** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideWidth**       | 以磅为单位返回或设置幻灯片的宽度。可读/写 **Single** 类型。  |

**成员属性**

#### **PageSetup.Application**

返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test(){   for(let i = 1 ; i <= Application.ActivePresentation.Slides.Item(1).Shapes.Count; i++){       if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject){           alert(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)       }   } }` |

#### **PageSetup.FirstSlideNumber**

返回或设置演示文稿中第一张幻灯片的编号。可读/写 **Long** 类型。

**语法**

**express.FirstSlideNumber**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

幻灯片编号是在显示幻灯片编号时出现在幻灯片右下角的实际编号。此编号是由演示文稿中幻灯片编号（顺序，即 **SlideIndex**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Slide/Slide%20.htm#Slide.SlideIndex)属性值）和演示文稿的起始幻灯片编号（**FirstSlideNumber** 属性值）所决定的。幻灯片编号始终等于起始幻灯片编号加上幻灯片索引号再减去一。**SlideNumber**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Slide/Slide%20.htm#Slide.SlideNumber)属性返回幻灯片编号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示更改活动演示文稿中第一张幻灯片的编号。*/ function test(){   let tion = Application.ActivePresentation   tion.PageSetup.FirstSlideNumber = 1         //starts slide numbering at 1   alert(tion.Slides.Item(2).SlideNumber)     //returns 2   tion.PageSetup.FirstSlideNumber = 10        //starts slide numbering at 10   alert(tion.Slides.Item(2).SlideNumber)     //returns 11 }` |

#### **PageSetup.NotesOrientation**

返回或设置指定演示文稿的备注页、讲义和大纲的屏幕显示和打印方向。可读/写 **MsoOrientation** 类型。

**语法**

**express.NotesOrientation**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

| MsoOrientation 可以是下列 MsoOrientation 常量之一。 |
| --------------------------------------------------- |
| **msoOrientationHorizontal**                        |
| **msoOrientationMixed**                             |
| **msoOrientationVertical**                          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将活动演示文稿中所有备注页、讲义和大纲的方向设为水平（横向）。 Application.ActivePresentation.PageSetup.NotesOrientation = msoOrientationHorizontal` |

#### **PageSetup.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。 function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **PageSetup.SlideHeight**

以磅为单位返回或设置幻灯片的高度。可读/写 **Single** 类型。

**语法**

**express.SlideHeight**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例会将当前演示文稿幻灯片的高度设为 8.5 英寸，宽度设为 11 英寸。*/ function test(){   let page = Application.ActivePresentation.PageSetup   page.SlideWidth = 11 * 72   page.SlideHeight = 8.5 * 72 }` |

#### **PageSetup.SlideOrientation**

返回或设置指定演示文稿中幻灯片的屏幕显示和打印方向。可读/写 **MsoOrientation** 类型。

**语法**

**express.SlideOrientation**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

| MsoOrientation 可以是下列 MsoOrientation 常量之一。 |
| --------------------------------------------------- |
| **msoOrientationHorizontal**                        |
| **msoOrientationMixed**                             |
| **msoOrientationVertical**                          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动演示文稿所有幻灯片的方向设为垂直（纵向）。*/ Application.ActivePresentation.PageSetup.SlideOrientation = msoOrientationVertical` |

#### **PageSetup.SlideSize**

返回或设置指定演示文稿的幻灯片大小。可读/写 **PpSlideSizeType** 类型。

**语法**

**express.SlideSize**

*express*   一个代表 **PageSetup** 对象的变量。

**说明**

| PpSlideSizeType 可以是下列 PpSlideSizeType 常量之一。 |
| ----------------------------------------------------- |
| **ppSlideSize35MM**                                   |
| **ppSlideSizeA3Paper**                                |
| **ppSlideSizeA4Paper**                                |
| **ppSlideSizeB4ISOPaper**                             |
| **ppSlideSizeB4JISPaper**                             |
| **ppSlideSizeB5ISOPaper**                             |
| **ppSlideSizeB5JISPaper**                             |
| **ppSlideSizeBanner**                                 |
| **ppSlideSizeCustom**                                 |
| **ppSlideSizeHagakiCard**                             |
| **ppSlideSizeLedgerPaper**                            |
| **ppSlideSizeLetterPaper**                            |
| **ppSlideSizeOnScreen**                               |
| **ppSlideSizeOverhead**                               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例将活动演示文稿的幻灯片设为投影机幻灯片大小。 Application.ActivePresentation.PageSetup.SlideSize = ppSlideSizeOverhead` |

#### **PageSetup.SlideWidth**

以磅为单位返回或设置幻灯片的宽度。可读/写 **Single** 类型。

**语法**

**express.SlideWidth**

*express*   一个代表 **PageSetup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例会将当前演示文稿幻灯片的高度设为 8.5 英寸，宽度设为 11 英寸。*/ function test(){   let page = Application.ActivePresentation.PageSetup   page.SlideWidth = 11 * 72   page.SlideHeight = 8.5 * 72 }` |

适用环境：web

适用平台：windows/linux