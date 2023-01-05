/*
    提示框
 */
function Message(option) {
    //这里做了一次初始化
    if(option == undefined) {
        return;
    }
    this.init(option);
}

Message.prototype.init = function (option) {
    //这里创建了提示框元素，并将整个提示框容器元素添加到页面中
    document.body.appendChild(this.create(option));
    //这里设置提示框的top
    this.setTop(document.querySelectorAll('.message'));
    //判断如果传入的closeTime大于0，则默认关闭提示框
    if (option.closeTime > 0) {
        this.close(option.container, option.closeTime);
    }
    //点击关闭按钮关闭提示框
    if (option.close) {
        option.close.onclick = (e) => {
            this.close(e.currentTarget.parentElement, 0);
        }
    }
}

Message.prototype.create = function (option) {
    //这里做了一个判断，表示如果设置showClose为false即不显示关闭按钮并且closeTime也为0，即无自动关闭提示框，我们就显示关闭按钮
    if(!option.showClose && option.closeTime <=0)option.showClose = true;
    //创建容器元素
    let element = document.createElement('div');
    //设置类名
    element.className = `message message-${option.type}`;
    //创建关闭按钮元素以及设置类名和内容
    let closeBtn = document.createElement('i');
    closeBtn.className = 'message-close-btn';
    closeBtn.innerHTML = '&times;';
    //创建内容元素
    let contentElement = document.createElement('p');
    contentElement.innerHTML = option.content;
    //判断如果显示关闭按钮，则将关闭按钮元素添加到提示框容器元素中
    if (closeBtn && option.showClose){}  element.appendChild(closeBtn);
    //将内容元素添加到提示框容器中
    element.appendChild(contentElement);
    //在配置项对象中存储提示框容器元素以及关闭按钮元素
    option.container = element;
    option.close = closeBtn;
    //返回提示框容器元素
    return element;
}

Message.prototype.close = function (messageElement, time) {
    //根据传入的时间来延迟关闭，实际上也就是移除元素
    setTimeout(() => {
    //判断如果传入了提示框容器元素，并且分两种情况，如果是多个提示框容器元素则循环遍历删除，如果是单个提示框容器元素，则直接删除
        if (messageElement && messageElement.length) {
            for (let i = 0; i < messageElement.length; i++) {
                if (messageElement[i].parentElement) {
                    messageElement[i].parentElement.removeChild(messageElement[i]);
                }
           }
        } else if (messageElement) {
            if (messageElement.parentElement) {
                 messageElement.parentElement.removeChild(messageElement);
            }
        }
        //关闭了提示框容器元素之后，我们重新设置提示框的top值
    this.setTop(document.querySelectorAll('.message'));
    }, time * 10);
}

Message.prototype.setTop = function (messageElement) {
    //这里做一个判断的原因就是当点击页面中最后一个提示框的时候，会重新调用一次，这时获取不到提示框容器元素，所以就不执行后续的设置top
    if (!messageElement || !messageElement.length) return;
    //由于每个提示框的高度一样，所以我们只需获取第一个提示框元素的高度即可
    const height = messageElement[0].offsetHeight;
    for (let i = 0; i < messageElement.length; i++) {
        //每个提示框的top由一个固定值加上它的高度，并且我们要乘以它的一个索引值
        messageElement[i].style.top = (25 * (i + 1) + height * i) + 'px';
    }
}

Message.prototype.info = function (content) {
    let option = {
        type: "info",
        closeTime: 200,
        showClose: true,
        content: content
    }
    //重新初始化
    this.init(option);
}

Message.prototype.warning = function (content) {
    let option = {
        type: "warning",
        closeTime: 200,
        showClose: true,
        content: content
    }
    //重新初始化
    this.init(option);
}
Message.prototype.success = function (content) {
    let option = {
        type: "success",
        closeTime: 200,
        showClose: true,
        content: content
    }
    //重新初始化
    this.init(option);
}

Message.prototype.error = function (content) {
    let option = {
        type: "error",
        closeTime: 200,
        showClose: true,
        content: content
    }
    //重新初始化
    this.init(option);
}

let $message = {}
window['$message'] = $message = new Message();
