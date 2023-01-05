/*
@ 负责API对象页面模板页面内容的填充，不与后端交互
*/
function FillObject(jsData) {
    if (jsData == undefined)
        return;
    document.title = jsData.name + " 对象"
    document.getElementById("template_header").innerHTML = document.title
    document.getElementById("template_header").id = "jsObject_" + jsData.name

    FillObjectDescription(jsData);     //添加summary
    FillObjectSummary(jsData);         //description
    document.getElementById("template_members").id = "jsMember_" + jsData.name; //成员列表标识符
    FillFunctionTable(jsData);         //table functions
    FillPropertyTable(jsData);         //table property
    FillFunctionsDetail(jsData);       //function...
    FillPropertiesDetail(jsData);      //property...

    if (typeof jsData.platform == "string" && jsData.platform != "") {
        document.getElementById("template_platform").innerHTML = jsData.platform
    }

    ScrollToElement(GetIdFromUrl())
}

/* 
    初始化对象说明
*/
function FillObjectDescription(jsData) {
    if(jsData.description === "") {
        return;
    }
    let template_description = document.getElementById("template_description");
    let p = document.createElement("p");
    p.innerHTML = jsData.description;
    template_description.appendChild(p);
}

/* 
    初始化对象详细说明
*/
function FillObjectSummary(jsData) {
    let template_summary = document.getElementById("template_summary");
    if(jsData.summary === "") {
        template_summary.innerHTML = "";
    }
    let p = document.createElement("p");
    p.innerHTML = jsData.summary;
    template_summary.appendChild(p)
    
}

/* 
    填充成员方法数据
*/
function FillFunctionTable(jsData) {
    if(jsData.functions == undefined) {
        return;
    }
    /* 成员区id */
    var tplFunctions = document.getElementById("member_functions");
    if(jsData.functions.length === 0) {
        tplFunctions.innerHTML = "";//如果方法为空，则不显示方法列表
        return;
    }
    
    var tplFuncs = tplFunctions.children[1].children[0].getElementsByTagName("tbody")[0];
    for (var idx in jsData.functions) {
        var fItem =jsData.functions[idx]
        var rowElem = document.createElement("tr");

        var url = "#" + jsData.name + "." + fItem.name
        rowElem.innerHTML = `
            <td width="6%" align="center" valian="middle"><img border=0 src="gif/methods.gif"></td>
            <td width="20%" align="center" valian="middle"><b class="bterm"><a href="${url}">${fItem.name}</a></b></td>
            <td style="word-break: break-all;" align="center" valian="middle">${fItem.description}</td>
        `
        rowElem.editName = "functions"; //设置编辑属性
        tplFuncs.appendChild(rowElem)
    }
}

/* 
    填充成员属性数据
*/
function FillPropertyTable(jsData) {
    if (jsData.properties == undefined)
        return
    var tplProperties = document.getElementById("member_properties");
    
    if(jsData.properties.length === 0) {
        tplProperties.innerHTML = "";//如果方法为空，则不显示方法列表
        return;
    }

    var tplProps = tplProperties.children[1].children[0].getElementsByTagName("tbody")[0];
    for (var idx in jsData.properties) {
        var pItem =jsData.properties[idx]
        var rowElem = document.createElement("tr");
        var url = "#" + jsData.name + "." + pItem.name
        rowElem.innerHTML = `
            <td width="6%" align="center" valian="middle"><img border=0 src="gif/properties.gif"></td>
            <td width="20%" align="center" valian="middle"><b class="bterm"><a href="${url}">${pItem.name}</a></b></td>
            <td style="word-break: break-all;" align="center" valian="middle">${pItem.description}</td>
        `
        rowElem.editName = "properties"; //设置编辑属性
        tplProps.appendChild(rowElem)
    }
}

/* 
    填充成员方法详情区
*/
function FillFunctionsDetail(jsData) {
    if (jsData.functions == undefined)
        return;
    
    let functions_details = document.getElementById("functions_details");
    if(jsData.functions.length === 0) {
        functions_details.innerHTML = "";//如果方法为空，则不显示方法列表
        return;
    }

    let tplDetailFuns = document.getElementById("funcs_details_content");
    for (var idx in jsData.functions) {
        var fItem =jsData.functions[idx]
        let elem = createMemDetailHtml("functions",jsData.name,fItem);
        tplDetailFuns.appendChild(elem);
    }
}


function FillPropertiesDetail(jsData) {
    if (jsData.properties == undefined)
        return;
    let properties_details = document.getElementById("properties_details");
    if(jsData.properties.length === 0) {
        properties_details.innerHTML = "";//如果方法为空，则不显示方法列表
        return;
    }

    var tplDetailProps = document.getElementById("props_details_content");
    for (var idx in jsData.properties) {
        var pItem = jsData.properties[idx];
        let elem = createMemDetailHtml("properties",jsData.name,pItem);
        tplDetailProps.appendChild(elem);
    }
}

/* 
    创建成员列表html元素
*/
function createMemDetailHtml(editName,objName,eleData) {
    if(eleData == undefined) {
        window.$message.error("参数eleData不能为空");
    }
    var itemId = objName + "." + eleData.name
    var elem = document.createElement("div");
    elem.id = itemId;
        
    //创建标题
    var title = document.createElement("h4");
    title.innerHTML = '<b>' + itemId + '</b>';
    title.style = "border-bottom: 2px #eee solid;padding: 15px 0 15px 0;margin: 1.5em 0 0.75em;"
    elem.appendChild(title)

    //创建主要内容区
    let elem_content = document.createElement("div");
    elem_content.style = "padding-left:40px";

    //语法
    let express = eleData.name;
    if(editName === "functions") {/* 如果时方法，有参数列表 */
        express += "(";
        for (var index = 0; index < eleData.parameters.length; ++index) {
            express += '<i>' + eleData.parameters[index].name + '</i>'
            if (index != eleData.parameters.length - 1)
                express += ", "
        }
        express += ")"
    }
    elem_content.innerHTML = `
        <p>${eleData.description}</p>
        <p><b class="mainheaders">语法</b></p>
        <p class="signature"><b><i STYLE="FONT-WEIGHT: normal">express.</i><span>${express}</span></b></p>
        <p><i STYLE="FONT-WEIGHT: normal">express</i>&nbsp;&nbsp;&nbsp;<span>一个代表 <b>${objName}</b> 对象的变量。</span></p>
    `

    //参数列表
    if (editName === "functions" && eleData.parameters.length > 0) {
        var pars = document.createElement("div");
        pars.innerHTML = '<p><b class="mainheaders">参数</b></p><div id="vstable"><table><tr><th><b>名称</b></th><th><b>必选/可选</b></th><th width="10%"><b>数据类型</b></th><th><b>说明</b></th></tr></table></div>'
        elem_content.appendChild(pars)

        var tplFuncs = pars.children[1].children[0]
        var paramTableText = '<td class="mainsection"><i>NAME</i></td><td class="mainsection">FLAG</td><td class="mainsection"><b>TYPE</b></td><td class="mainsection"  style="word-break: break-all;" >TEXT</td>'
        let paramLen = eleData.parameters.length;
        for (var idx = 0;idx < paramLen;++idx) {
            var pItem = eleData.parameters[idx]
            var elem_tr = document.createElement("tr");
            elem_tr.innerHTML = paramTableText.replace("NAME", pItem.name).replace("FLAG", pItem.optional?"可选":"必选").replace("TYPE", pItem.type).replace("TEXT", pItem.description);
            tplFuncs.appendChild(elem_tr)
        }
    }

    //添加返回值或者成员属性类型
    if((eleData.returns && eleData.returns !== "") || (eleData.type && eleData.type !== "")){
        let lan = document.createElement("p");
        lan.innerHTML = '<b class="mainheaders">' + (editName === "functions" ? "返回值" : "类型") +'</b>';
        elem_content.appendChild(lan)
        var p = document.createElement("p");
        p.innerHTML = editName === "functions" ? eleData.returns : eleData.type;
        elem_content.appendChild(p)
    }

    //添加方法说明
    if(eleData.summary && eleData.summary !== "") {
        var elem_summary = document.createElement("p");
        elem_summary.innerHTML = '<b class="mainheaders">说明</b>';
        elem_content.appendChild(elem_summary)
        
        var tplExplain = document.createElement("div")
        tplExplain.innerHTML = '<p>' + eleData.summary + '</p>';
        elem_content.appendChild(tplExplain);
    }
    /* 
        添加示例代码
    */
   if(eleData.examples && eleData.examples !== "") {
        var elem_examples = document.createElement("p");
        elem_examples.innerHTML = '<b class="mainheaders">示例</b>';
        elem_content.appendChild(elem_examples);

        var eleCode = document.createElement("div");
        eleCode.innerHTML = '<p>' + eleData.examples + '</p>';
        elem_content.appendChild(eleCode);    
   }

   elem.appendChild(elem_content);
   return elem;
}

function GetIdFromUrl() {
    var url = decodeURI(location.href)
    var nameList = url.split("#")
    if (nameList.length !== 0) {
       return nameList[nameList.length - 1]
    }
    return "#";
}

function ScrollToElement(name) {
    var ele = document.getElementById(name)
    window.scrollTo({
        top:heightToTop(ele),
        behavior:'auto'
    })
}

function heightToTop(ele){
    //ele为指定跳转到该位置的DOM节点
    let bridge = ele;
    let root = document.body;
    let height = 0;
    do{
        height += bridge.offsetTop;
        bridge = bridge.offsetParent;
    }while(bridge !== root)
 
    return height;
}

function InitData(wpsData, cbFillData) {
    objectData = wpsData
    objectExtData.refLink = []
    cbFillData(objectData)
}