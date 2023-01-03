**ActionSettings**



包含两个 **ActionSetting** 对象的集合，这些对象用于形状或文本范围。

**说明**

其中一个 **ActionSetting** 对象代表在幻灯片放映中用户单击指定对象时该对象的反应；另一个 **ActionSetting** 对象代表在幻灯片放映中用户将鼠标指针移动到指定对象上时该对象的反应。

**方法**

|                                                              | 名称     | 说明                                             |
| ------------------------------------------------------------ | -------- | ------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从指定的 ActionSettings 集合中返回单个动作设置。 |

**属性**

|                                                              | 名称            | 说明                                                    |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 Application 对象，该对象表示指定对象的创建者。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回指定集合中的对象数目。                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。                                  |

**成员方法**

#### **ActionSettings.Item**

从指定的 ActionSettings 集合中返回单个动作设置。

**语法**

**express.Item(Index)**

*express*   一个代表 **ActionSettings** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**          | **说明**                                 |
| -------- | ------------- | --------------------- | ---------------------------------------- |
| *Index*  | 必选          | **PpMouseActivation** | MouseClick 或 MouseOver 事件的动作设置。 |

**返回值**

ActionSetting

**说明**

| PpMouseActivation 可为以下 PpMouseActivation 常数之一。 |
| ------------------------------------------------------- |
| ppMouseClick 用户单击形状时的动作设置。                 |
| ppMouseOver 当鼠标指针定位在指定形状之上时的动作设置。  |

**Item** 方法是集合的默认成员。

有关返回集合中单个成员的详细信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置第一张幻灯片中的第三个形状播放掌声，并使用 AnimateAction 属性指定在幻灯片放映中用鼠标单击该形状时，形状的颜色立即反转。*/ function test(){ 　　　　let Act = Application.ActivePresentation.Slides.Item(1).Shapes.Item(3).ActionSettings.Item(ppMouseClick) 　　　　Act.SoundEffect.Name = "applause" 　　　　Act.AnimateAction = true }` |

**成员属性**

#### **ActionSettings.Application**

返回一个 Application 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **ActionSettings** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，一个 Presentation 对象被传递至某过程。此过程在演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function AddAndSave(pptPres) {     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test(){ 　　　　for(let i=1;i <= Application.ActivePresentation.Slides.Item(1).Shapes.Count;i++) {    　　　　 if(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).Type == msoLinkedOLEObject) {       　　　　  alert(Application.ActivePresentation.Slides.Item(1).Shapes.Item(i).OLEFormat.Application.Name)    　　　　 } 　　　　} }` |

#### **ActionSettings.Count**

返回指定集合中的对象数目。

**语法**

**express.Count**

*express*   一个代表 **ActionSettings** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭除窗口1外的所有窗口。*/ function test(){ 　　　　let Win = Application.Windows     for(let i=2;i <= Win.Count;i++) { 　　　　    Win.Item(2).Close() 　　　　} }` |

#### **ActionSettings.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ActionSettings** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。*/ function test(){ 　　　　let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes 　　　　let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame 　　　　addShp.TextRange.Text = "Test text" 　　　　addShp.Parent.Rotation = 45 }` |

适用环境：web

适用平台：windows/linux