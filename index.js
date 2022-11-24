function doPost(e) {
    var channelToken = "";
    var value = JSON.parse(e.postData.contents);
    try {
      var SpreadSheet = SpreadsheetApp.openById("1wAL-Nn6zq4ERP1zZQRQILWQcdwBsNHQh5vwz9H3Yov8");
      var Sheet = SpreadSheet.getSheets()[0];
      var events = value.events;
      if (events != null) {
        for (var i in events) {
          var event = events[i];
          var type = event.type;
          var replyToken = event.replyToken; //
          var userId = event.source.userId; // 取得個人userId
          var groupId = event.source.groupId; // 取得群組Id
          var LastRow = Sheet.getLastRow();
          Sheet.getRange(LastRow + 1, 1).setValue(new Date());
          Sheet.getRange(LastRow + 1, 2).setValue(userId);
          Sheet.getRange(LastRow + 1, 3).setValue(event.message.text);
          Sheet.getRange(LastRow + 1, 4).setValue(replyToken);
          Sheet.getRange(LastRow + 1, 5).setValue(groupId);
          switch (type) {
            case "message":
              var a = event.message.text;
              a = a.replace("","_").replace("“","%E2%80%9C").replace("”","%E2%80%9D").replace("\xE2\x80\x8E","%E2%80%A7").replace("‧","%E2%80%A7").replace("\\","%22");
              if ((a.indexOf("[[") > -1 && a.indexOf("]]") > -1) || (a.indexOf("{{") > -1 && a.indexOf("}}") > -1)) {
                var response = [];
                var list = a.split("[[");
                for (var i = 1; i < list.length; i++) {
                  var text = list[i];
                  response.push("https://zh.m.wikipedia.org/wiki/" + text.split("]]")[0]);
                }
                list = a.split("{{");
                for (i = 1; i < list.length; i++) {
                  text = list[i];
                  response.push("https://zh.m.wikipedia.org/wiki/Template:" + text.split("}}")[0]);
                }
                meg = [[
                  {
                    "type": "flex",
                    "altText": "Search REsult",
                    "contents": {
                      "type": "bubble",
                      "size": "kilo",
                      "body": {
                        "type": "box",
                        "layout": "vertical",
                        "contents": [
                          {
                            "type": "text",
                            "text": "EM Linkbot",
                            "weight": "bold",
                            "color": "#1DB446",
                            "size": "sm",
                            "margin": "none"
                          },
                          {
                            "type": "text",
                            "text": "Brown Store",
                            "weight": "bold",
                            "size": "xxl",
                            "margin": "none",
                            "adjustMode": "shrink-to-fit"
                          },
                          {
                            "type": "text",
                            "text": "hello, world"
                          },
                          {
                            "type": "box",
                            "layout": "horizontal",
                            "contents": [
                              {
                                "type": "image",
                                "url": "https://raw.githubusercontent.com/Edit-Mr/EM-Wikibot/main/static/wiki.png",
                                "size": "30px",
                                "flex": 1,
                                "action": {
                                  "type": "uri",
                                  "label": "action",
                                  "uri": "http://linecorp.com/"
                                }
                              },
                              {
                                "type": "image",
                                "url": "https://raw.githubusercontent.com/Edit-Mr/EM-Wikibot/main/static/google.png",
                                "size": "30px",
                                "action": {
                                  "type": "uri",
                                  "label": "action",
                                  "uri": "http://linecorp.com/"
                                },
                                "flex": 1
                              },
                              {
                                "type": "image",
                                "url": "https://raw.githubusercontent.com/Edit-Mr/EM-Wikibot/main/static/youtube.png",
                                "size": "30px",
                                "flex": 1,
                                "action": {
                                  "type": "uri",
                                  "label": "action",
                                  "uri": "http://linecorp.com/"
                                }
                              }
                            ],
                            "justifyContent": "flex-start",
                            "alignItems": "flex-start"
                          }
                        ]
                      },
                      "styles": {
                        "footer": {
                          "separator": true
                        }
                      }
                    }
                  }
                ]];
                Sheet.getRange(LastRow + 1, 6).setValue(meg);
                replyMsg(replyToken, meg, channelToken);
              } else Sheet.getRange(LastRow + 1, 6).setValue("No reply");
              break;
            default:
              Sheet.getRange(LastRow + 1, 6).setValue("Not Text");
              break;
          }
  
        }
      }
    } catch (ex) {
      Sheet.getRange(LastRow + 1, 7).setValue(ex);
    }
  }
  // 回覆訊息
  function replyMsg(replyToken, userMsg, channelToken) {
    try {
      var url = "https://api.line.me/v2/bot/message/reply";
      var opt = {
        headers: {
          "Content-Type": "application/json; charset=UTF-8",
          Authorization: "Bearer " + channelToken,
        },
        method: "post",
        payload: JSON.stringify({
          replyToken: replyToken,
          muteHttpExceptions: true,
          messages: userMsg,
        }),
      };
      UrlFetchApp.fetch(url, opt);
    } catch (ex) {
      Sheet.getRange(LastRow + 1, 7).setValue(ex);
    }
  }