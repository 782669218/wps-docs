**TableStyle**



代表可应用于表格或切片器的单个样式。

**说明**

表格样式为表格、数据透视表或切片器的一个或所有元素定义格式。例如，列是表格的元素。表格样式可以规定使用交替格式（也称为条带或条纹）对表格中的列进行格式设置。

**方法**

|                                                              | 名称          | 说明                                                 |
| ------------------------------------------------------------ | ------------- | ---------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**    | 删除 **TableStyle** 对象。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Duplicate** | 复制 **TableStyle** 对象，并返回对新复制对象的引用。 |

**属性**

|                                                              | 名称                               | 说明                                                         |
| ------------------------------------------------------------ | ---------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                    | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltIn**                        | 如果样式为内置样式，则为 **True**。只读 **Boolean** 类型。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                        | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                           | 返回对象的名称。只读 **String** 类型。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NameLocal**                      | 以用户语言返回或设置对象的名称。只读 **String** 类型。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                         | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAsAvailablePivotTableStyle** | 设置或返回一个值，该值表示是否在数据透视表样式库中显示样式，可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAsAvailableSlicerStyle**     | 返回或设置指定的表样式是否作为可用的表样式在切片器样式库中显示。可读/写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowAsAvailableTableStyle**      | 返回或设置在表样式库中显示为可用的表样式。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TableStyleElements**             | 返回 **TableStyleElements** 对象，只读。                     |

**成员方法**

#### **TableStyle.Delete**

删除 **TableStyle** 对象。

**语法**

**express.Delete()**

*express*   一个代表 **TableStyle** 对象的变量。

**返回值**

无

#### **TableStyle.Duplicate**

复制 **TableStyle** 对象，并返回对新复制对象的引用。

**语法**

**express.Duplicate(NewTableStyleName)**

*express*   一个代表 **TableStyle** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**         |
| ------------------- | ------------- | ------------ | ---------------- |
| *NewTableStyleName* | 必选          | **Variant**  | 新表样式的名称。 |

**返回值**

TableStyle

**说明**

如果未指定名称，则 **Duplicate** 方法将使用与用户界面默认名称相同的默认名称。

**成员属性**

#### **TableStyle.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **TableStyle** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **TableStyle.BuiltIn**

如果样式为内置样式，则为 **True**。只读 **Boolean** 类型。

**语法**

**express.BuiltIn**

*express*   一个代表 **TableStyle** 对象的变量。

#### **TableStyle.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **TableStyle** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **TableStyle.Name**

返回对象的名称。只读 **String** 类型。 

**语法**

**express.Name**

*express*   一个代表 **TableStyle** 对象的变量。

#### **TableStyle.NameLocal**

以用户语言返回或设置对象的名称。只读 **String** 类型。

**语法**

**express.NameLocal**

*express*   一个代表 **TableStyle** 对象的变量。

**说明**

如果样式为内置样式，则此属性以当前系统所使用的地区语言返回样式的名称。

#### **TableStyle.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **TableStyle** 对象的变量。

#### **TableStyle.ShowAsAvailablePivotTableStyle**

设置或返回一个值，该值表示是否在数据透视表样式库中显示样式，可读/写 **Boolean** 类型。

**语法**

**express.ShowAsAvailablePivotTableStyle**

*express*   一个代表 **TableStyle** 对象的变量。

**说明**

如果在数据透视表样式库中显示样式，则返回 **True**。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 即使已将样式应用于表格或数据透视表，用户仍可以将 **ShowAsAvailableTableStyle** 或 **ShowAsAvailablePivotTableStyle** 属性设置为 **False**。在这种情况下，当活动单元格位于表格或数据透视表中时，库将不会显示样式并且没有样式显示为选中的样式。 |

#### **TableStyle.ShowAsAvailableSlicerStyle**

返回或设置指定的表样式是否作为可用的表样式在切片器样式库中显示。可读/写

**语法**

**express.ShowAsAvailableSlicerStyle**

*express*   一个代表 **TableStyle** 对象的变量。

#### **TableStyle.ShowAsAvailableTableStyle**

返回或设置在表样式库中显示为可用的表样式。可读/写 **Boolean** 类型。

**语法**

**express.ShowAsAvailableTableStyle**

*express*   一个代表 **TableStyle** 对象的变量。

**说明**

如果为 **True**，则此样式显示在表样式库中。

即使样式已应用于表，也可以将此属性设置为 **False**。这种情况下，库中将不会显示该样式，活动单元格在该表中时，样式状态不会显示为选中。

#### **TableStyle.TableStyleElements**

返回 **TableStyleElements** 对象，只读。

**语法**

**express.TableStyleElements**

*express*   一个代表 **TableStyle** 对象的变量。

适用环境：web

适用平台：windows/linux