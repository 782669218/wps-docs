**AddIns**



**AddIn** 对象的集合，这些对象代表与 WPP 相关的对WPP 可用的所有加载项，而无论这些加载项是否已加载。

**说明**

这里不包含组件对象模型 (COM) 加载项 （COM 加载项：通过添加自定义命令和指定的功能来扩展 WPS Office程序的功能的补充程序。COM 加载项可在一个或多个 Office 程序中运行。COM 加载项使用文件扩展名 .dll 或 .exe。）。

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**    | 返回一个 AddIn 对象，该对象代表一个添加到加载宏列表中的加载宏文件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 从指定集合中返回单个对象。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Remove** | 从加载宏集合中删除一个加载宏。                               |

**属性**

|                                                              | 名称            | 说明                                                    |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                  |

**成员方法**

#### **AddIns.Add**

返回一个 AddIn 对象，该对象代表一个添加到加载宏列表中的加载宏文件。

**语法**

**express.Add(FileName)**

*express*   一个代表 **AddIns** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName* | 必选          | **String**   | 包含要添加到加载宏列表中的加载宏的文件的完整名称（包括路径和文件扩展名）。 |

**返回值**

AddIn

**说明**

此方法并不加载新的加载宏。必须设置 **Loaded** 属性才能加载该加载宏。

#### **AddIns.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **AddIns** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

AddIn

#### **AddIns.Remove**

从加载宏集合中删除一个加载宏。

**语法**

**express.Remove(Index)**

*express*   一个代表 **AddIns** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                       |
| -------- | ------------- | ------------ | ------------------------------ |
| *Index*  | 必选          | **Variant**  | 要从集合中删除的加载宏的名称。 |

**示例**

本示例从可用加载宏的列表中删除名为“MyTools”的加载宏。

| 示例代码复制                                     |
| ------------------------------------------------ |
| `示例代码  Appliaction.AddIns.Remove("mytools")` |

**成员属性**

#### **AddIns.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **AddIns** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test(){         for(let i=1;i <= Application.ActivePresentation.Slides.Item(1).Shapes.Count;i++) {             if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {              　　　alert(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)           　}         }     }` |

#### **AddIns.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **AddIns** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除窗口1外的所有窗口。*/ function test(){ 　　　　let Win = Application.Windows        for(let i=2;i <= Win.Count;i++) { 　　　　    Win.Item(2).Close() 　　　　} }` |

#### **AddIns.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **AddIns** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux