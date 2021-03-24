/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import axios from 'axios';

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */
    
    const body = context.document.body;

    // Queue a command to load the text in document body.
    context.load(body, 'text');

    await context.sync();

    const url = "https://wordaddin.cognitiveservices.azure.com/text/analytics/v3.1-preview.1/entities/recognition/general";
    const headers = {
      'Content-Type': 'application/json',
      'Ocp-Apim-Subscription-Key': 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
    }
    const data = {
      "documents": [
        {
          "language": "en",
          "id": 1,
          "text": body.text
        }
      ]
    }
    var entities = [];
    var allSearchResults = [];
    const sendPostRequest = async () => {
      try {
        const response = await axios.post(url, data, {headers: headers});
        var doc = response.data.documents[0];
        entities = doc.entities;
      } catch(error) {
        console.log(error);
      }
    };
    await sendPostRequest();

    var entity_map = {};
    for (let i = 0; i < entities.length; i++){
      if(entity_map[entities[i].category]){
        entity_map[entities[i].category] = entity_map[entities[i].category]+", "+String(entities[i].text);
      }
      else{
        entity_map[entities[i].category] = String(entities[i].text);
      }
    }
    var para = "";
    for (var key in entity_map){
      para = para + "<b>"+key+"</b>" + " - " + entity_map[key] + "<br />";
    }
    //var allSearchResults = [];
    for (let i = 0; i < entities.length; i++){
     // para = para + entities[i].category + " - " + entities[i].text + "<br />";
      var searchResults = context.document.body.search(entities[i].text, {ignorePunct: true});
      // Queue a command to load the search results and get the font property values.
      //searchResults.load("text");
      context.load(searchResults, 'font');
      allSearchResults.push(searchResults);
    }
    document.getElementById("entities").innerHTML = para;

    await context.sync();

    for (let k = 0; k < allSearchResults.length; k++) {
      var curResult = allSearchResults[k];
      //console.log(curResult.items);
      for (let j = 0; j < curResult.items.length; j++) {
        console.log(curResult.items[j].font.color);
        curResult.items[j].font.color = 'red';
        curResult.items[j].font.highlightColor = '#00FF00'; //Yellow
        curResult.items[j].font.bold = true;
      }
      //await context.sync();
    }

    await context.sync();
    // // insert a paragraph at the end of the document.
    // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
    // // change the paragraph color to blue.
    // paragraph.font.color = "blue";
    // await context.sync();
  })
  .catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  });
}
