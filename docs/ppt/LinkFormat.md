**LinkFormat**



包含应用于链接 OLE 对象的属性和方法。

**说明**

**OLEFormat** 对象包含应用于链接和非链接 OLE 对象的属性和方法。**PictureFormat** 对象包含应用于图片和 OLE 对象的属性和方法。

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Update** | 更新指定的链接 OLE 对象。要一次更新演示文稿中的所有链接，可使用 **UpdateLinks** 方法。 |

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**    | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoUpdate**     | 返回或设置更新链接的方法。可读/写 **PpUpdateOption** 类型。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**         | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceFullName** | 返回或设置链接 OLE 对象的源文件的名称和路径。可读/写 **String** 类型。 |

**成员方法**

#### **LinkFormat.Update**

更新指定的链接 OLE 对象。要一次更新演示文稿中的所有链接，可使用 **UpdateLinks** 方法。

**语法**

**express.Update()**

*express*   一个代表 **LinkFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更新当前演示文稿中所有链接的 OLE 对象。*/ function test(){ 　　　　for(let i = 1; i <= Application.ActivePresentation.Slides.Count; i++){     　　　　for(let ii = 1; ii <= Application.ActivePresentation.Slides.Item(i).Shapes.Count; ii++){        　　　　 if(Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).Type == msoLinkedOLEObject){ 　　　　            Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).Type.LinkFormat.Update() 　　　　        } 　　　　    } 　　　　}}` |

**成员属性**

#### **LinkFormat.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **LinkFormat** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **LinkFormat.AutoUpdate**

返回或设置更新链接的方法。可读/写 **PpUpdateOption** 类型。

**语法**

**express.AutoUpdate**

*express*   一个代表 **LinkFormat** 对象的变量。

**说明**

| PpUpdateOption 可以是下列 PpUpdateOption 常量之一。          |
| ------------------------------------------------------------ |
| ppUpdateOptionAutomatic 每次打开演示文稿或更改源文件时更新链接。 |
| ppUpdateOptionManual 仅当用户特别要求更新演示文稿时更新链接。 |
| ppUpdateOptionMixed                                          |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例浏览活动演示文稿中所有幻灯片的所有形状，并将所有链接的 ET 工作表设为手动更新。*/ function test(){ 　　　　for(let i = 1; i<= Application.ActivePresentation.Slides.Count; i++){   　　　　  for(let ii = 1; ii <= Application.ActivePresentation.Slides.Item(i).Shapes.Count; ii++){ 　　　　        if(Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).Type == msoLinkedOLEObject){ 　　　　            if(Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).OLEFormat.ProgID == "Excel.Sheet"){  　　　　               Application.ActivePresentation.Slides.Item(i).Shapes.Item(ii).LinkFormat.AutoUpdate = ppUpdateOptionManual 　　　　            }  　　　　       }  　　　　   } 　　　　} }` |

#### **LinkFormat.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **LinkFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes 　　　　let tex = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　tex.TextRange.Text = "Test text" 　　　　tex.Parent.Rotation = 45 }` |

#### **LinkFormat.SourceFullName**

返回或设置链接 OLE 对象的源文件的名称和路径。可读/写 **String** 类型。

**语法**

**express.SourceFullName**

*express*   一个代表 **LinkFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前演示文稿第一张幻灯片第一个形状的源文件设为 Wordtest.doc，并指定该对象的图像自动更新。*/ function test(){ 　　　　let sha = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1)    　　　　 if(sha.Type == msoLinkedOLEObject){     　　　　    let lin = sha.LinkFormat     　　　　    lin.SourceFullName = "C:\\my documents\\wordtest.doc"      　　　　   lin.AutoUpdate = ppUpdateOptionAutomatic   　　　　 } }` |

适用环境：web

适用平台：windows/linux