**NamedSlideShow**



代表自定义幻灯片放映，该幻灯片放映是演示文稿中幻灯片的一个命名子集。

**说明**

**NamedSlideShow** 对象是 **NamedSlideShows** 集合的成员。**NamedSlideShows** 集合包含演示文稿中的所有命名幻灯片放映。

**方法**

|                                                              | 名称       | 说明             |
| ------------------------------------------------------------ | ---------- | ---------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 不能使用该属性设置自定义幻灯片放映的名称。使用 **Add** 方法可以新名称再定义一个自定义幻灯片放映。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideIDs**    | 返回指定的命名幻灯片放映的幻灯片 ID 数组。只读 **Variant** 类型。 |

**成员方法**

#### **NamedSlideShow.Delete**

删除指定的对象。

**语法**

**express.Delete()**

*express*   一个代表 **NamedSlideShow** 对象的变量。

**说明**

如果试图删除表格中唯一一个现有行或列，将导致运行时错误。

**成员属性**

#### **NamedSlideShow.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **NamedSlideShow** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **NamedSlideShow.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **NamedSlideShow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除当前窗口外的所有窗口。*/ function test(){ 　　　　let win = Application.Windows    　　　　 for(let i = 2; i <= win.Count; i++){     　　　　    win.Item(i).Close()  　　　　   } }` |

#### **NamedSlideShow.Name**

不能使用该属性设置自定义幻灯片放映的名称。使用 **Add** 方法可以新名称再定义一个自定义幻灯片放映。**String** 类型，只读。

**语法**

**express.Name**

*express*   一个代表 **NamedSlideShow** 对象的变量。

**说明**

如果包含某对象的集合的 **Item** 方法采用 **Variant** 类型的参数，则可以将该对象的名称与 **Item** 方法联合使用来返回对该对象的引用。例如，如果某个形状的 **Name** 属性值为“Rectangle 2”，则 `.Shapes("Rectangle 2")` 将返回对该形状的引用。

#### **NamedSlideShow.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **NamedSlideShow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes 　　　　let sha = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　sha.TextRange.Text = "Test text" 　　　　sha.Parent.Rotation = 45 }` |

#### **NamedSlideShow.SlideIDs**

返回指定的命名幻灯片放映的幻灯片 ID 数组。只读 **Variant** 类型。

**语法**

**express.SlideIDs**

*express*   一个代表 **NamedSlideShow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口中的当前幻灯片添加到名为“Marketing Short Version”的自定义幻灯片放映中。请注意，要保存修改后的自定义幻灯片放映版本，必须删除原始自定义放映，再使用相同名称重新添加。另请注意，要调整包含于 Variant 变量中的数组大小，必须在调整该数组大小前明确声明该变量。*/ function test(){ //NOTE - The following code line is NOT optional. //Can't redim array without this 									 　　　　customShowName = "Marketing Short Version" 　　　　let customShowToExpand = Application.ActivePresentation.SlideShowSettings.NamedSlideShows.Item(customShowName) 　　　　let slideToAddID = Application.ActiveWindow.View.Slide.SlideID 　　　　let customShowSlideIDs = customShowToExpand.SlideIDs 　　　　let numSlides = UBound(customShowSlideIDs) 									 　　　　customShowSlideIDs(numSlides + 1) = slideToAddID 　　　　customShowToExpand.Delete() 　　　　ActivePresentation.SlideShowSettings.NamedSlideShows.Add(customShowName, customShowSlideIDs) }` |

适用环境：web

适用平台：windows/linux