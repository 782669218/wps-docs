**OLEFormat**



包含应用于 OLE 对象的属性和方法。

**说明**

**LinkFormat** 对象包含仅应用于链接 OLE 对象的属性和方法。**PictureFormat** 对象包含应用于图片和 OLE 对象的属性和方法。

**方法**

|                                                              | 名称         | 说明                                                         |
| ------------------------------------------------------------ | ------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate** | 激活指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DoVerb**   | 请求 OLE 对象执行其中一个动作。使用 **ObjectVerbs** 属性可确定 OLE 对象的可用动作。 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FollowColors** | 返回或设置指定对象中的颜色对幻灯片配色方案的继承程度。指定的对象必须是在 Graph 或 Organization Chart 中创建的图表。可读/写 **PpFollowColors** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Object**       | 返回一个对象，该对象代表指定 OLE 对象的顶层界面。该属性使您可以访问某些应用程序中的属性和方法（OLE 对象在这些应用程序中创建）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ObjectVerbs**  | 返回一个 **ObjectVerbs** 集合，该集合包含指定的 OLE 对象的所有 OLE动作。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProgID**       | 返回指定 OLE 对象的程序标识符 (ProgID) （程序标识符：窗体 OLEServerName.ObjectName 中的标识符（例如，Excel.Sheet 或WPP.Slide），由 Windows 注册表用来唯一标识一个对象。）。只读 **String** 类型。 |

**成员方法**

#### **OLEFormat.Activate**

激活指定的对象。

**语法**

**express.Activate()**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

本示例按顺序激活紧挨在活动窗口之后的文档窗口。

| 示例代码复制                             |
| ---------------------------------------- |
| `Application.Windows.Item(2).Activate()` |

#### **OLEFormat.DoVerb**

请求 OLE 对象执行其中一个动作。使用 **ObjectVerbs** 属性可确定 OLE 对象的可用动作。

**语法**

**express.DoVerb(Index)**

*express*   一个代表 **OLEFormat** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                       |
| -------- | ------------- | ------------ | ---------------------------------------------- |
| *Index*  | 必选          | **Integer**  | 要执行的动作。如果忽略该参数，则执行默认动作。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对当前演示文稿中第一张幻灯片的第三个形状执行默认动作，其中第三个形状是链接或嵌人的 OLE 对象。*/ function test(){ 　　　　let sli = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)  　　　　   if(sli.Type == msoEmbeddedOLEObject || sli.Type == msoLinkedOLEObject){      　　　　   sli.OLEFormat.DoVerb()   　　　　  } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对当前演示文稿中第一张幻灯片的第三个形状执行默认动作“Open”，其中第三个形状是支持动作“Open”的 OLE 对象。* / function test(){ 　　　　let sli = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3)   　　　　  if(sli.Type == msoEmbeddedOLEObject || sli.Type == msoLinkedOLEObject){     　　　　    for(let i = 1; i <= sli.OLEFormat.ObjectVerbs.Count; i++){       　　　　      nCount++        　　　　     if(sli.OLEFormat.ObjectVerbs.Item(i) == "Open"){         　　　　        sli.OLEFormat.DoVerb(nCount)       　　　　      }    　　　　     } 　　　　   } }` |

**成员属性**

#### **OLEFormat.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPres){     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **OLEFormat.FollowColors**

返回或设置指定对象中的颜色对幻灯片配色方案的继承程度。指定的对象必须是在 Graph 或 Organization Chart 中创建的图表。可读/写 **PpFollowColors** 类型。

**语法**

**express.FollowColors**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

| PpFollowColors 可以是下列 PpFollowColors 常量之一。          |
| ------------------------------------------------------------ |
| ppFollowColorsNone 图表颜色不继承幻灯片配色方案。            |
| ppFollowColorsMixed                                          |
| ppFollowColorsScheme 图表中所有的颜色都继承幻灯片配色方案。  |
| ppFollowColorsTextAndBackground 只有文本和背景继承幻灯片配色方案。 |

**示例**

本示例指定当前演示文稿第一张幻灯片第二个形状的文本和背景继承幻灯片配色方案。第二个形状必须是 Graph 或 Organization Chart 创建的图表。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.Slides.Item(1).Shapes.Item(2).OLEFormat.FollowColors = ppFollowColorsTextAndBackground` |

#### **OLEFormat.Object**

返回一个对象，该对象代表指定 OLE 对象的顶层界面。该属性使您可以访问某些应用程序中的属性和方法（OLE 对象在这些应用程序中创建）。只读。

**语法**

**express.Object**

*express*   一个代表 **OLEFormat** 对象的变量。

**说明**

使用 **TypeName** 函数判断此属性为特定的 OLE 对象返回的对象的类型。

**示例**

本示例显示当前演示文稿第一张幻灯片第一个形状中包含的对象的类型。第一个形状必须包含 OLE 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `alert(TypeName(Application.ActivePresentation.Slides.Item(1).Shapes.Item(1).OLEFormat.Object))` |

#### **OLEFormat.ObjectVerbs**

返回一个 **ObjectVerbs** 集合，该集合包含指定的 OLE 对象的所有 OLE动作。只读。

**语法**

**express.ObjectVerbs**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示了当前演示文稿中第二张幻灯片上第一个形状内的 OLE 对象具有的所有可用操作。只有当第一个形状是 OLE 对象时，这个例子才会有效。*/ function test(){ let ole = Application.ActivePresentation.Slides.Item(2).Shapes.Item(1).OLEFormat     for(let i = 1; i <= ole.ObjectVerbs.Count; i++){         alert(ole.ObjectVerbs)     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例表明：对于当前演示文稿上第二张幻灯片中第一个形状所代表的 OLE 对象而言，若“Open”为该对象的一个 OLE 操作，那么在幻灯片演示过程中如果单击该对象，则此对象会打开。只有当第一个形状是 OLE 对象时，这个例子才会有效。*/ function test() {     let sli = Application.ActivePresentation.Slides.Item(2).Shapes.Item(1)     for (let i = 1; i <= sli.OLEFormat.ObjectVerbs.Count; i++) {         if (sli.OLEFormat.ObjectVerbs.Item(i) == "Open") {             let set = sli.ActionSettings(ppMouseClick)             set.Action = ppActionOLEVerb             set.ActionVerb = sli.OLEFormat.ObjectVerbs.Item(i)         }     } }` |

#### **OLEFormat.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes 　　　　let tex = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　tex.TextRange.Text = "Test text" 　　　　tex.Parent.Rotation = 45 }` |

#### **OLEFormat.ProgID**

返回指定 OLE 对象的程序标识符 (ProgID) （程序标识符：窗体 OLEServerName.ObjectName 中的标识符（例如，Excel.Sheet 或WPP.Slide），由 Windows 注册表用来唯一标识一个对象。）。只读 **String** 类型。

**语法**

**express.ProgID**

*express*   一个代表 **OLEFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例搜索当前演示文稿所有幻灯片中的所有对象，将全部链接的 ET 工作表设为手动更新。*/ function test() {     let te = Application.ActivePresentation.Slides     for (let i = 1; i <= te.Count; i++) {         for (let ii = 1; ii <= te.Item(i).Shapes;) {             if (te.Item(i).Shapes.Item(ii).Type == msoLinkedOLEObject) {                 if (te.Item(i).Shapes.Item(ii).OLEFormat.ProgID == "Excel.Sheet") {                     te.Item(i).Shapes.Item(ii).LinkFormat.AutoUpdate = ppUpdateOptionManual                 }             }         }     } }` |

适用环境：web

适用平台：windows/linux