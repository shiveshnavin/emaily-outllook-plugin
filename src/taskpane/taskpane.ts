import axios from 'axios'
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */
  const running = document.getElementById("run")
  running.innerText = 'Running...'
  const item = Office.context.mailbox.item;

  const body: Office.AsyncResult<string> = await new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, resolve)
  })
  let data = JSON.stringify({
    "text": `Tell me action items in short pointers from below email msg, output ur repose as html: ${body.value}`,
    "options": {
      "conversationId": "c_078e7ece0d0d1137"
    }
  });

  let config = {
    method: 'post',
    maxBodyLength: Infinity,
    url: '/query',
    headers: {
      'accept': 'application/json, text/plain, */*',
      'accept-language': 'en-US,en;q=0.9,de;q=0.8',
      'cache-control': 'no-cache',
      'content-type': 'application/json',
      'cookie': '_ga=GA1.1.274938184.1694205355; _ga_R1FN4KJKJH=GS1.1.1703357946.3.1.1703358770.0.0.0; _ga_MH88ELNX5E=GS1.1.1714573613.17.1.1714573660.0.0.0',
      'pragma': 'no-cache',
      'priority': 'u=1, i',
      'sec-ch-ua': '"Not)A;Brand";v="99", "Google Chrome";v="127", "Chromium";v="127"',
      'sec-ch-ua-mobile': '?0',
      'sec-ch-ua-platform': '"Windows"',
      'sec-fetch-dest': 'empty',
      'sec-fetch-mode': 'cors',
      'sec-fetch-site': 'same-origin',
      'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36'
    },
    data: data
  };

  let q = await axios.request(config)
    .then((response) => {
      console.log(JSON.stringify(response.data));
      return (response.data.response)
    })
    .catch((error) => {
      console.log(error);
      return JSON.stringify(error?.response?.data) || error.message
    });

  let insertAt = document.getElementById("item-subject");
  insertAt.appendChild(document.createElement("br"));
  var div = document.createElement('div');
  div.innerHTML = `<div style='width:250px;'>${q}</div>`;
  insertAt.appendChild(div);
  insertAt.appendChild(document.createElement("br"));

  running.innerText = ''
}
