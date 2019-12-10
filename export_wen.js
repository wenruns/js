let export_wen = function ({
                               elementID = null,
                               url = '',
                               data = {},
                               fileName = '',
                               excelHead = [],
                               excelBody = [],
                               method = 'GET',
                               buttonStyle = '',
                               tipsMsg = '正在导出',
                               fontFamily = "'Source Sans Pro','Helvetica Neue',Helvetica,Arial,sans-serif",
                               headerFunc = null,
                               footerFunc = null,
                               dataHandle = null,
                               dataFormat = null,
                           }) {
    let element_id = null;
    let excelHtml = '';
    let excelData = [];
    let page = 0;
    let cancelExport = false;

    this.tipsElement = null;

    this.wenAjax = function ({
                                 url = '',
                                 method = 'GET',
                                 header = {},
                                 async = true,
                                 data = {},
                                 success = null,
                                 fail = null,
                                 completed = null,
                             }) {
        function createClient() {
            if (XMLHttpRequest) {
                return new XMLHttpRequest();
            } else if (ActiveXObject) {
                // 支持IE7之前的版本
                if (typeof arguments.callee.activeXString !== 'string') {
                    var versions = ['MSXML2.XMLHttp.6.0', 'MSXML2.XMLHttp.3.0', 'MSXML2.XMLHttp'];
                    for (var i = 0; i < versions.length; i++) {
                        try {
                            new ActiveXObject(versions[i]);
                            arguments.callee.activeXString = versions[i];
                            break;
                        } catch (e) {
                            //
                        }
                    }
                }
                return new ActiveXObject(arguments.callee.activeXString);
            } else {
                throw new Error("No XHR Object available!");
            }
        }

        let xhr = createClient();
        xhr.onreadystatechange = function () {
            if (xhr.readyState == 4) {
                let completeResponse = xhr;
                if ((xhr.status >= 200 && xhr.status < 300) || xhr.status == 304) {
                    completeResponse = xhr.responseText;
                    success && success(xhr.responseText);
                } else {
                    fail && fail(xhr)
                }
                completed && completed(completeResponse);
            }
        }

        let paramStr = '';
        for (var key in data) {
            if (paramStr) {
                paramStr += "&" + encodeURIComponent(key) + "=" + encodeURIComponent(data[key])
            } else {
                paramStr += encodeURIComponent(key) + "=" + encodeURIComponent(data[key])
            }
        }
        switch (method) {
            case 'GET':
                if (url.indexOf('?') > 0) {
                    url += '&' + paramStr;
                } else {
                    url += '?' + paramStr;
                }
                paramStr = '';
                break;
            default:
        }
        xhr.open(method, url, async); // open()方法启动一个请求以备发送；
        if (method == 'POST') {
            xhr.setRequestHeader('content-type', 'application/x-www-form-urlencoded;charset=UTF-8');
        }
        if (header.length) {
            for (var index in header) {
                xhr.setRequestHeader(index, header[index]);
            }
        }
        xhr.send(paramStr);
    }

    function __init() {
        if (!elementID) {
            console.error('元素id（elementID）不能为空或者null，请输入元素id');
            return false;
        }
        this._element = new Array();
        bindEvent();
    }

    let bindEvent = () => {
        if (typeof elementID == 'string') {
            elementID = [elementID];
        }
        elementID.forEach(function (item, dex) {
            try {
                this._element[dex] = document.querySelector('#' + item);
                if (!this._element[dex]) {
                    throw new Error("'#" + item + "' is not a valid selector.");
                }
            } catch (e) {
                console.error(e)
                return false;
            }
            this._element[dex].addEventListener('click', function (e) {
                if (element_id) {
                    console.error("the export is runing by the id='#" + element_id + "'");
                    return false;
                }
                if (!e.target.dataset.configure) {
                    console.error("The attribute 'data-configure' is not defined on the '#" + item + "'");
                    return false;
                }
                let config = new Array();
                try {
                    config = JSON.parse(e.target.dataset.configure)
                } catch (err) {
                    console.error(err);
                    return false;
                }
                if (config.url) {
                    url = config.url;
                }
                if (config.method) {
                    method = config.method;
                }
                if (config.excelBody) {
                    excelBody = config.excelBody;
                }
                if (config.excelHead) {
                    excelHead = config.excelHead;
                }
                if (config.fileName) {
                    fileName = config.fileName;
                }
                if (config.data) {
                    Object.assign(data, config.data)
                }
                if (!url) {
                    console.error('请求路径（url）不能为空或者null');
                    return false;
                }
                if (!excelBody.length) {
                    console.error('导出字段信息（excelBody）不能为空');
                    return false;
                }
                if (!excelHead.length) {
                    console.error('导出头信息（excelHead）不能为空');
                    return false;
                }
                if (!fileName) {
                    if (document.title) {
                        fileName = document.title + '.xls'
                    } else {
                        let dateObj = new Date();
                        fileName = dateObj.getFullYear() + '' + (dateObj.getMonth() + 1) + '' + dateObj.getDate() + '' + dateObj.getHours() + '' + dateObj.getMinutes() + dateObj.getSeconds() + '_YmdHis.xls';
                    }
                }
                element_id = item;
                showTips();
                initData();
                getData();
            })
        })
    }


    function initData() {
        excelHtml = '';
        page = 0;
        excelData = [];
        cancelExport = false;
    }

    let showTips = () => {
        this.tipsElement = createElement();
        this.tipsElement.querySelector('.wen-cancel-button').addEventListener('click', function (e) {
            hideTips();
            cancelExport = true;
        })
        document.body.appendChild(this.tipsElement);
    }

    let hideTips = () => {
        if (this.tipsElement) {
            document.body.removeChild(this.tipsElement);
        }
        element_id = null;
    }

    function styleParse(style) {
        if (typeof style == 'string') {
            return style;
        }
        let styleStr = '';
        for (var attr in style) {
            styleStr += attr + ':' + style[attr] + ';';
        }
        return styleStr;
    }

    function createElement() {
        let divObj = document.createElement('div');
        divObj.innerHTML = '<div>' +
            '   <style>' +
            '       .wen-change{' +
            '            width: 40%;' +
            '            margin-top: 15%;' +
            '            -webkit-animation: mymove 2s infinite;' +
            '            animation: mymove 2s infinite;' +
            '        }' +
            '        @keyframes mymove' +
            '        {' +
            '            0% {-webkit-transform:rotate(0deg);}' +
            '            50% {-webkit-transform:rotate(180deg);}' +
            '            100% {-webkit-transform:rotate(360deg);}' +
            '        }' +
            '        @-webkit-keyframes mymove /*Safari and Chrome*/' +
            '        {' +
            '            0% {-webkit-transform:rotate(0deg);}' +
            '            50% {-webkit-transform:rotate(180deg);}' +
            '            100% {-webkit-transform:rotate(360deg);}' +
            '        }' +
            '   </style>' +
            '    <div style="width: 100vw; height: 100vh; background: black; opacity: 0.5; position: fixed; top: 0px; left: 0px; z-index: 100;"></div>' +
            '    <div style="width: 100vw; height: 100vh; position: fixed; z-index: 101; left: 0vw; top: 0vh; border: 1px solid red;">' +
            '        <div style="width: 200px; height: 200px; background: white; border-radius: 5px; margin: 30vh auto;">' +
            '            <div style="height: 75%;text-align: center;">' +
            '                <img class="wen-change" src="/images/static/loading.png" alt="">' +
            '                <div>' + tipsMsg + '</div>' +
            '            </div>' +
            '            <div style="text-align: center;">' +
            '                <button style="cursor: pointer;padding: 3px 15px;' + styleParse(buttonStyle) + '" class="wen-cancel-button">取消</button>' +
            '            </div>' +
            '        </div>   ' +
            '    </div>' +
            '</div>';
        return divObj;
    }


    // 生成下载文件
    function download() {
        makeHead();
        makeBody();
        makeFooter();
        // 生成excel表格
        let blob = new Blob([excelHtml], {
            type: 'application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        if (window.navigator.msSaveOrOpenBlob) {
            navigator.msSaveBlob(blob);
        } else {
            let elink = document.createElement('a');
            elink.download = fileName;
            elink.style.display = 'none';
            elink.href = URL.createObjectURL(blob);
            document.body.appendChild(elink);
            elink.click();
            document.body.removeChild(elink);
        }
        hideTips();
    }

    function makeHead() {
        excelHtml = '<' + 'html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
            'xmlns:x="urn:schemas-microsoft-com:office:excel" ' +
            'xmlns="http://www.w3.org/TR/REC-html40"><head>' +
            '<!--[if gte mso 9]>' +
            '<xml>' +
            '<x:ExcelWorkbook>' +
            '   <x:ExcelWorksheets>' +
            '       <x:ExcelWorksheet>' +
            '           <x:Name><' + '/' + 'x:Name>' +
            '           <x:WorksheetOptions>' +
            '               <x:DisplayGridlines' + '/' + '>' +
            '           <' + '/' + 'x:WorksheetOptions>' +
            '       <' + '/' + 'x:ExcelWorksheet>' +
            '   <' + '/' + 'x:ExcelWorksheets>' +
            '<' + '/' + 'x:ExcelWorkbook>' +
            '<' + '/' + 'xml>' +
            '<![endif]-->' +
            '<style>' +
            "   br{mso-data-placement:same-cell;}" +
            "   .wen-excel-header{font-family: " + fontFamily + ";}" +
            "   .wen-excel-head{text-align: center;height: 34px; font-family:" + fontFamily + ";}" +
            "   .wen-excel-body{text-align: center;height: 28px; font-family:" + fontFamily + ";}" +
            "   .wen-excel-statistic{text-align: left !important;height: 28px; font-family:" + fontFamily + ";}" +
            '<' + '/style>' +
            '</head><body><table><thead>';
        if (headerFunc) {
            let str = headerFunc(excelData, element_id);
            if (str) {
                excelHtml += checkString(str);
            }
        }
        excelHtml += '<tr>';
        excelHead.forEach(function (item, dex) {
            excelHtml += '<th class="wen-excel-head">' + item + '</th>'
        })
        excelHtml += '</tr></thead><tbody>';
    }


    function makeBody() {
        excelData.forEach(function (item, dex) {
            for (var i = 0; i < item.length; i++) {
                excelHtml += '<tr>';
                excelBody.forEach(function (field, index) {
                    excelHtml += '<td class="wen-excel-body">' + format(item[i], field) + '</td>';
                })
                for (var field in excelBody) {
                }
                excelHtml += '</tr>';
            }
        })
    }

    function format(data, field) {
        if (dataFormat) {
            return dataFormat(data, field, element_id)
        }
        return data[field];
    }

// 生成脚部
    function makeFooter() {
        if (footerFunc) {
            let str = footerFunc(excelData, element_id);
            if (str) {
                excelHtml += checkString(str);
            }
        }
        excelHtml += '<' + '/tbody><' + '/table><' + '/body><' + '/html>';
    }

// 头部和脚部输出额外的字符串（多行请使用数组）
    function checkString(str) {
        let arr = str;
        if (typeof arr != 'object' && !(arr instanceof Array)) {
            arr = new Array();
            arr.push(str);
        }
        let len = excelHead.length;
        let str1 = '';
        let tr = '';
        let i = 0;
        for (var dex in arr) {
            str1 += arr[dex] + '<br/>';
            if (i > 0) {
                tr += '<tr></tr>';
            }
            i++;
        }
        var reg = new RegExp('(<br/>$)', 'gi');
        str1 = str1.replace(reg, "");
        let htmlStr = '<tr class="wen-excel-header"><td rowspan="' + arr.length + '" class="xl68" height="' + (arr.length * 28) + '" colspan="' + len + '" style="height: ' + (arr.length * 28) + 'px;border-right:none;border-bottom:none;" x:str>' + str1 + '</td></tr>' + tr;
        return htmlStr;
    }

    let getData = () => {
        console.log('cancelExport=' + cancelExport);
        if (cancelExport) {
            cancelExport = false;
            return false;
        }
        data['page'] = page;
        // 发送ajax请求
        this.wenAjax({
            url: url,
            method: method,
            data: data,
            success: function (res) {
                if (dataHandle) {
                    res = dataHandle(res, element_id);
                }
                if (typeof res == 'string') {
                    res = JSON.parse(res);
                }
                if (res.length > 0 || Object.keys(res).length) {
                    excelData.push(res);
                    page++;
                    getData();
                } else {
                    // 导出excel
                    download();
                }
            },
            fail: function (err) {
                console.error(err)
                // return false;
            }
        });
    }
    __init();
}