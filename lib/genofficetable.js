var EMU = 914400; // OfficeXML measures in English Metric Units


module.exports = {

  // assume passed in an array of row objects
  getTable: function(rows, options) {
    var options = options || {};
    options.tabstyle=options.tabstyle?options.tabstyle:"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
    if (options.columnWidth === undefined) {
      options.columnWidth = 8 / (rows[0].length) * EMU
    }
    var self = this;

    return self._getBase(
      rows.map(function(row,row_idx) {
        return self._getRow(
          row.map(function(val,idx) {
            var cellVal = val, cellOptions = options;
            if ((typeof val === 'object')) { //Cell-specific formatting passed in, override table options
              cellOptions = (val.hasOwnProperty('opts')) ? val.opts : options;
              cellOptions.font_size = cellOptions.font_size || options.font_size
              cellOptions.font_face = cellOptions.font_face || options.font_face
              cellOptions.align = cellOptions.align || options.align
              cellOptions.fill_color = cellOptions.fill_color || options.fill_color
              cellOptions.font_color = cellOptions.font_color || options.font_color
              cellOptions.margin = cellOptions.margin || options.margin || {}
              cellVal = (val.hasOwnProperty('val')) ? val.val : val;
            }

            return self._getCell(cellVal, cellOptions, options,idx,row_idx);
          }),
          row_idx,
          options
        );
      }),
      self._getColSpecs(rows, options),
      options
    )
  },

  "_getBase": function (rowSpecs, colSpecs, options) {
    var self = this;

    return {
      "p:graphicFrame": {
        "p:nvGraphicFramePr": {
          "p:cNvPr": {
            "@id": "6",
            "@name": "Table 5",
            "a:extLst": {
              "a:ext": {
                "@uri": "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}",
                "a16:creationId": {
                  "@xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main",
                  "@id": "{6B19EF81-4905-48BF-8149-71011CE43ED6}"
                }
              }
            }
          },
          "p:cNvGraphicFramePr": {
            "a:graphicFrameLocks": {
              "@noGrp": "1"
            }
          },
          "p:nvPr": {
            "p:extLst": {
              "p:ext": {
                "@uri": "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}",
                "p14:modId": {
                  "@xmlns:p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
                  "@val": "1579011935"
                }
              }
            }
          }
        },
        "p:xfrm": {
          "a:off": {
            "@x": options.x || "1524000",
            "@y": options.y || "1397000"
          },
          "a:ext": {
            "@cx": options.cx || "6096000",
            "@cy": options.cy || "1483360"
          }
        },
        "a:graphic": {
          "a:graphicData": {
            "@uri": "http://schemas.openxmlformats.org/drawingml/2006/table",
            "a:tbl": {
              "a:tblPr": {
                "@firstRow": "1",
                "@bandRow": "1",
                "a:tableStyleId":options.tabstyle//"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
              },
              "a:tblGrid": {
                "#list": colSpecs
              },

              "#list": [rowSpecs]  // replace this with  an array of table row objects
            }
          }
        }
      }
    }
  },

  _getColSpecs: function(rows, options) {
    var self = this;
    return rows[0].map(function(val,idx) {
      return self._tblGrid(idx, options);
    })
  },

  _tblGrid: function(idx, options) {
    return {
      "a:gridCol": {
        "@w": (options.columnWidths ? options.columnWidths[idx] : options.columnWidth|| "2048000" ),
        "a:extLst": {
          "a:ext": {
            "@uri": "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}",
            "a16:colId": {
              "@xmlns:a16": "http://schemas.microsoft.com/office/drawing/2014/main",
              "@val": "1074158453"
            }
          }
        }
      }
    };
  },

  _getRow: function (cells, row_idx, options) {
    return {
      "a:tr": {
        "@h": (options.rowHeights ? options.rowHeights[row_idx] : options.rowHeight || "0" ), //|| "370840"
        "#list": [cells] // populate this with an array of table cell objects
      }
    }
  },

  _getCell: function (val, options,idx,row_idx) {
    if(options.vMerge || options.hMerge){
      return {
        "a:tc": {
          "@vMerge": options.vMerge ? '1' : '0',
          "@hMerge": options.hMerge ? '1' : '0',
          "a:txBody": {
            "a:bodyPr": {},
            "a:lstStyle": {},
            "a:p": {
              "a:endParaRpr": {
                "@lang": "zh-CN",
                "altLang": "en-US",
                "dirty": "0"
              }
            }
          },
        }
      }
    }
    var font_size = options.font_size || 14;
    var font_face = options.font_face || "Times New Roman";
    var cellObject = {
      "a:tc": {
        "@rowSpan": options.row_span || 1,
        "@gridSpan": options.grid_span || 1,
        "a:txBody": {
          "a:bodyPr": {},
          "a:lstStyle": {},

          "a:p": {
            "a:pPr": {"@algn": options.align ? (options.align[idx] ? options.align[idx] : options.align) : 'ctr'},
            "a:r": {
              "a:rPr": {
                "@lang": "en-US",
                "@sz": "" + font_size * 100,
                "@dirty": "0",
                "@smtClean": "0",
                "@b": options.bold ? (options.bold[row_idx] ? (options.bold[row_idx][idx] ? options.bold[row_idx][idx] : options.bold[row_idx] ) : options.bold) : "0",
                "@i": options.italics ? (options.italics[row_idx] ? (options.italics[row_idx][idx] ? options.italics[row_idx][idx] : options.italics[row_idx]) : options.italics) : "0",
                "a:latin": {
                  "@typeface": font_face
                },
                "a:cs": {
                  "@typeface": font_face
                },
                "a:solidFill": {
                  "a:srgbClr": {
                    "@val": options.font_color || "000000",
                  }
                },
              },
              "a:t": val  // this is the cell value
            },
            "a:endParaRPr": {
              "@lang": "en-US",
              "@sz": "" + font_size * 100,
              "@dirty": "0",
              "a:solidFill": {
                "a:srgbClr": {
                  "@val": options.font_color || "000000",
                }
              },
              "a:latin": {
                "@typeface": font_face
              },
              "a:cs": {
                "@typeface": font_face
              }
            }
          }
        },
        "a:tcPr": {
          "@anchor": 'ctr',
          "@marL": options.margin.left || "85233",
          "@marR": options.margin.right || "85233",
          "@marT": options.margin.top || "42617",
          "@marB": options.margin.bottom || "42617",
          "a:solidFill": {
            "a:srgbClr": {
              "@val": options.fill_color || "ffffff",
              "a:alpha": {
                "@val": (100 - (options.fill_color_alpha || 0)) * 1000
              }
            }
          }
        }
      }
    };

    return cellObject;
  }
}