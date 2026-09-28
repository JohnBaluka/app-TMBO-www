window.domWrapper = {
    GetElementById: function (id) {
        return document.getElementById(id);
    },
    SetElementText: function (id, text) {
        const element = document.getElementById(id);
        if (element) {
            element.textContent = text;
        }
    },
    SetElementStyle: function (id, styleProperty, value) {
        const element = document.getElementById(id);
        if (element) {
            element.style[styleProperty] = value;
        }
    },
    GetElementsByClassName: function (className) {
        return Array.from(document.getElementsByClassName(className)).map(el => el.outerHTML);
    }
};