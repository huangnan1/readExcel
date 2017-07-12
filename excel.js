var character = new Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z");
var $getId = function(_this) {
    var result = _this;
    if (typeof _this == 'object') {
        result = eval('_this.id');
        if (result == undefined) result = eval('_this.Id');
    }
    return result;
};

Array.prototype.find = function(funCompareOrValue, start) {
    var result = null;
    if (this) {
        for (var i = 0; i < this.length; i++) {
            if (start && start > i) continue;
            if (typeof funCompareOrValue == 'function' ? funCompareOrValue(this[i], i) : $getId(this[i]) == funCompareOrValue) {
                result = this[i];
                break;
            }
        }
    }
    return result;
};
Array.prototype.finds = function(funCompareOrValue) {
    var results = [];
    if (this) {
        for (var i = 0; i < this.length; i++) {
            if ((typeof funCompareOrValue == 'function' ? funCompareOrValue(this[i], i) : $getId(this[i]) == funCompareOrValue) && results.indexOf(this[i]) == -1) {
                results.push(this[i]);
            }
        }
    }
    return results;
};

function readExcel(files) {
    var file = files[0];
    var reader = new FileReader();
    var name = file.name;
    reader.onload = function(e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
            type: 'binary'
        });
        if (workbook.SheetNames.length) {
            var btns = document.querySelector('.select');
            workbook.SheetNames.forEach(function(item) {
                var input = document.createElement('input');
                input.type = 'button';
                input.value = item;
                btns.appendChild(input);
            });
            btns.onclick = function(e) {
                var e = e || window.event;
                var target = e.target || e.srcElement;
                if (target.nodeName.toLowerCase() == 'input') {
                    getExcelToJson(workbook.Sheets[target.value]);


                }

            };
        }
    };
    reader.readAsBinaryString(file);
};

function getExcelToJson(sheet, colConvertList) {
    delete sheet["!ref"];
    var json = {};
    for (var key in sheet) {
        var rowIndex = key.substring(1) - 1;
        var colIndex = key.substring(0, 1); //colConvertList[key.substring(0, 1)];
        if (!json[rowIndex]) json[rowIndex] = {};
        json[rowIndex][colIndex] = sheet[key].w;
    }
    var tempAry = [];
    for (var index in json) {
        tempAry[index] = json[index];
    }
    if (colConvertList) console.log(JSON.stringify(tempAry));
    else createTable(tempAry, sheet['!merges']);
};

function createTable(tempAry, merges) {
    var wrap = document.querySelector('.excelData');
    wrap.innerHTML = '';

    var table = document.createElement('table');
    var tdsLength = [];
    tempAry.forEach(function(obj) {
        var propertys = Object.getOwnPropertyNames(obj);
        tdsLength.push(character.indexOf(propertys[propertys.length - 1]));
    });
    tempAry.forEach(function(item, index) {
        var tr = document.createElement('tr');
        for (var i = 0; i <= Math.max.apply(Math, tdsLength); i++) {
            var td = document.createElement(!index ? 'th' : 'td');
            td.innerHTML = item[character[i]] ? '<span style="color:red;">' + character[i] + '</span>' + item[character[i]] : '';
            td.setAttribute('_id', i);
            tr.appendChild(td);
        }
        table.appendChild(tr);
    });
    Array.prototype.forEach.call(table.childNodes, function(item, index) {
        var curRowMerges = merges.finds(function(m) {
            return m.s.r == index;
        });
        curRowMerges.forEach(function(coordinate) {
            var colspan = coordinate.e.c - coordinate.s.c;
            var rowspan = coordinate.e.r - coordinate.s.r;

            if (rowspan) {
                rowspan++;
                Array.prototype.find.call(item.childNodes, function(f) {
                    return f.getAttribute('_id') == [coordinate.s.c];
                }).rowSpan = rowspan;
                for (var k = 1; k < rowspan; k++) {
                    var curRows = table.childNodes[coordinate.s.r + k];
                    curRows.removeChild(
                        Array.prototype.find.call(curRows.childNodes, function(f) {
                            return f.getAttribute('_id') == coordinate.s.c;
                        })
                    );
                }

            }
            if (colspan) {
                colspan++;
                Array.prototype.find.call(item.childNodes, function(f) {
                    return f.getAttribute('_id') == [coordinate.s.c];
                }).colSpan = colspan;
                for (var r = 1; r < colspan; r++) {
                    item.removeChild(
                        Array.prototype.find.call(item.childNodes, function(f) {
                            return f.getAttribute('_id') == (coordinate.s.c + r);
                        }));
                }
                if (rowspan) {
                    for (var k = 1; k < rowspan; k++) {
                        var curRows = table.childNodes[coordinate.s.r + k];
                        for (var r = 1; r < colspan; r++) {
                            curRows.removeChild(
                                Array.prototype.find.call(curRows.childNodes, function(f) {
                                    return f.getAttribute('_id') == (coordinate.s.c + r);
                                }));
                        }
                    }
                }
            }
        });
    });
    wrap.appendChild(table);
}