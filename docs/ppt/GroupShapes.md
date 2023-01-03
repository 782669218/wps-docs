**GroupShapes**



代表组合形状中的各个形状。每个形状都由一个 **Shape** 对象代表。对本对象使用 **Item** 方法，可以在不取消分组的情况下处理组合中的各个形状。

**方法**

|                                                              | 名称      | 说明                           |
| ------------------------------------------------------------ | --------- | ------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**  | 从指定集合中返回单个对象。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Range** | 返回一个 **ShapeRange** 对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                       |

**成员方法**

#### **GroupShapes.Item**

从指定集合中返回单个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **GroupShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                               |
| -------- | ------------- | ------------ | -------------------------------------- |
| *Index*  | 必选          | **Variant**  | 集合中要返回的单个对象的名称或索引号。 |

**返回值**

Shape

#### **GroupShapes.Range**

返回一个 **ShapeRange** 对象。

**语法**

**express.Range(Index)**

*express*   一个代表 **GroupShapes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要包含在范围中的各个形状。可以是指定形状的索引号的 Integer、指定形状名称的 String，或者是包含整数或字符串的数组。如果省略此参数，则 Range 方法将返回指定集合中的所有对象。 |

**返回值**

ShapeRange

**说明**

虽然可以使用 **Range** 方法返回任意数目的形状或幻灯片，但如果只需要返回该集合的一个成员，则使用 **Item** 方法会更为简单。例如，`Shapes(1)` 比 `Shapes.Range(1)` 简单，`Slides(2)` 比 `Slides.Range(2)` 简单。

若要为 **Index** 指定一个整数或字符串数组，可以使用 **Array** 函数。例如，以下指令返回用名称指定的两个形状。

`let myArray = ["Oval 4", "Rectangle 5"] let myRange = ActivePresentation.Slides.Item(1).Shapes.Range(myArray)`

**成员属性**

#### **GroupShapes.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **GroupShapes** 对象的变量。

**示例**

以下示例中，一个 **Presentation** 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function AddAndSave(pptPre) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes     for (let i = 1; i <= shpOle.Count; i++) {         if (shpOle.Item(i).Type == msoLinkedOLEObject) {             alert(shpOle.Item(i).OLEFormat.Application.Name)         }     } }` |

#### **GroupShapes.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **GroupShapes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除窗口1外的所有窗口。*/ function test(){ 　　　　for(let i=2; i<=Application.Windows.Count; i++) {  　　　　   Application.Windows.Item(i).Close() 　　　　} }` |

#### **GroupShapes.Creator**

返回 **Long** 类型值，该值代表创建指定对象的应用程序创建者代码，该代码由四个字符构成。例如，如果对象是在WPP 中创建的，则此属性返回一个十六进制数 50575054。只读。

**语法**

**express.Creator**

*express*   一个代表 **GroupShapes** 对象的变量。

**说明**

**Creator** 属性用于 Macintosh 下的 WPS Office应用程序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示关于 myObject 创建者的消息。*/ function test(){ 　　　　let myObject = Application.ActivePresentation.Slides.Item(1).Shapes.Item(1) 　　　　if (myObject.Creator == "0x50575054") { 　　　　    alert("This is aWPP object") 　　　　} 　　　　else { 　　　　    alert("This is not aWPP object") 　　　　} }` |

#### **GroupShapes.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **GroupShapes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux