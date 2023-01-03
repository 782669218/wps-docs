**SlideShowWindows**



代表WPP 中打开的幻灯片放映的所有 [SlideShowWindow](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/SlideShowWindow/SlideShowWindow%20.htm#jsObject_SlideShowWindow) 对象的集合。

**方法**

|                                                              | 名称     | 说明                             |
| ------------------------------------------------------------ | -------- | -------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 集合中要返回的单个对象的索引号。 |

**属性**

|                                                              | 名称            | 说明                                                    |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                              |

**成员方法**

#### **SlideShowWindows.Item**

集合中要返回的单个对象的索引号。

**语法**

**express.Item(Index)**

*express*   一个代表 **SlideShowWindows** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                         |
| -------- | ------------- | ------------ | -------------------------------- |
| *Index*  | 必选          | **Long**     | 集合中要返回的单个对象的索引号。 |

**成员属性**

#### **SlideShowWindows.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **SlideShowWindows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中*/ function test(){     let pptPres = Application.ActivePresentation     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide")  }   /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称*/ function test(){     let shpOle = ActivePresentation.Slides.Item(1).Shapes     for(let i = 1; i <= shpOle.Count; i++) {         if(shpOle.Item(i).Type == msoLinkedOLEObject) {             MsgBox(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **SlideShowWindows.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **SlideShowWindows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除窗口1以外的所有窗口*/ function test(){     for(let i = 2; i <= Application.Windows.Count; i++) {         Application.Windows.Item(2).Close()     } }` |

适用环境：web

适用平台：windows/linux