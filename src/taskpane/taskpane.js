/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { async } from "regenerator-runtime";

/* global document, Office, Word */

Office.onReady((info) => {
  console.log(info);
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("editConfig").onclick = editConfig;
    document.getElementById("saveConfig").onclick = saveConfig;
    document.getElementById("publish").onclick = publish;
    document.getElementById("preview").onclick = preview;


  }

  checkConfig();
  // var fileUrl = Office.context.document.url;
  // console.log(fileUrl);

});

export async function run() {
  return Word.run(async (context) => {
    checkConfig();

    await context.sync();
  });
}

export async function getInitialState() {
  return Word.run(async (context) => {
    var fileUrl = Office.context.document.url;
    var aemRepo = Office.context.document.settings.get('aemRepo');
    var aemRepoName = aemRepo.replace('https://github.com/', '');
    var productionUrl = Office.context.document.settings.get('productionUrl');
    var contentUrl = Office.context.document.settings.get('contentUrl');
    var previewButton = document.getElementById("preview");
    var publishButton = document.getElementById("publish");
    var pageMetadata = document.getElementById("pageMetadata");
    var fileUrl = Office.context.document.url;

    // convert spaces to %20 in fileUrl
    fileUrl = fileUrl.replace(' ', '%20');

    // strip contentUrl from fileUrl
    var fileUrl = fileUrl.replace(contentUrl, '');

    // convert spaces to - in fileUrl
    fileUrl = fileUrl.replace(/ /g, '-');

    // remove .docx from fileUrl
    fileUrl = fileUrl.replace('.docx', '');

    // remove ’ from fileUrl
    fileUrl = fileUrl.replace(/’/g, '-');

    // convert fileUrl to lowercase
    fileUrl = fileUrl.toLowerCase();

    var liveUrl = 'https://admin.hlx.page/status/' + aemRepoName + fileUrl;
    console.log(liveUrl);
    fetch(liveUrl, {
      method: "GET",
    })
      .then((response) => response.json())
      .then((json) => {

        console.log(json);

        // find element with id lastModified
        var lastModified = document.getElementById('lastModified');
        lastModified.innerHTML = `Last modified: ${json.preview.lastModified}`;

        // get iframe
        var iframe = document.getElementById('aemPage');
        // reload iframe with preview url
        iframe.src = `${json.preview.url}?date=${Date.now()}`;
        iframe.addEventListener('load', handleLoad, true)

        function handleLoad() {
          console.log('iframe loaded');
          loader.classList.add('d-none');
          previewButton.textContent = "Preview";
          publishButton.textContent = "Publish";
          pageMetadata.classList.remove('d-none');
        }


      });
  });
}



export async function preview() {

  return Word.run(async (context) => {
    var fileUrl = Office.context.document.url;
    var aemRepo = Office.context.document.settings.get('aemRepo');
    var aemRepoName = aemRepo.replace('https://github.com/', '');
    var productionUrl = Office.context.document.settings.get('productionUrl');
    var contentUrl = Office.context.document.settings.get('contentUrl');
    var previewButton = document.getElementById("preview");
    previewButton.textContent = "Previewing...";

    var loader = document.getElementById("loader");
    loader.classList.remove('d-none');

    // convert spaces to %20 in fileUrl
    fileUrl = fileUrl.replace(' ', '%20');

    // strip contentUrl from fileUrl
    var fileUrl = fileUrl.replace(contentUrl, '');

    // convert spaces to - in fileUrl
    fileUrl = fileUrl.replace(/ /g, '-');

    // remove .docx from fileUrl
    fileUrl = fileUrl.replace('.docx', '');

    // remove ’ from fileUrl
    fileUrl = fileUrl.replace(/’/g, '-');

    // convert fileUrl to lowercase
    fileUrl = fileUrl.toLowerCase();

    var liveUrl = 'https://admin.hlx.page/preview/' + aemRepoName + fileUrl;
    console.log(liveUrl);
    fetch(liveUrl, {
      method: "POST",
    })
      .then((response) => response.json())
      .then((json) => {

        console.log(json);

        // find element with id lastModified
        var lastModified = document.getElementById('lastModified');
        lastModified.innerHTML = `Last modified: ${json.preview.lastModified}`;

        function handleLoad() {
          console.log('iframe loaded');
          loader.classList.add('d-none');
          previewButton.textContent = "Preview";
        }

        // get iframe
        var iframe = document.getElementById('aemPage');
        // reload iframe with preview url
        iframe.src = `${json.preview.url}?date=${Date.now()}`;
        iframe.addEventListener('load', handleLoad, true)
      });


    await context.sync();

  });
}


export async function publish() {
  return Word.run(async (context) => {
    var fileUrl = Office.context.document.url;
    var aemRepo = Office.context.document.settings.get('aemRepo');
    var aemRepoName = aemRepo.replace('https://github.com/', '');
    var productionUrl = Office.context.document.settings.get('productionUrl');
    var contentUrl = Office.context.document.settings.get('contentUrl');
    var publishButton = document.getElementById("publish");
    publishButton.textContent = "Publishing...";

    var loader = document.getElementById("loader");
    loader.classList.remove('d-none');


    // convert spaces to %20 in fileUrl
    fileUrl = fileUrl.replace(' ', '%20');

    // strip contentUrl from fileUrl
    var fileUrl = fileUrl.replace(contentUrl, '');

    // convert spaces to - in fileUrl
    fileUrl = fileUrl.replace(/ /g, '-');

    // remove .docx from fileUrl
    fileUrl = fileUrl.replace('.docx', '');

    // remove ’ from fileUrl
    fileUrl = fileUrl.replace(/’/g, '-');

    // convert fileUrl to lowercase
    fileUrl = fileUrl.toLowerCase();

    var liveUrl = 'https://admin.hlx.page/live/' + aemRepoName + fileUrl;
    console.log(liveUrl);
    fetch(liveUrl, {
      method: "POST",
      body: null,
    })
      .then((response) => response.json())
      .then((json) => {
        console.log(json);

        // find element with id lastModified
        var lastModified = document.getElementById('lastModified');
        lastModified.innerHTML = `Last modified: ${json.live.lastModified}`;

        function handleLoad() {
          console.log('iframe loaded');
          loader.classList.add('d-none');
          publishButton.textContent = "Publish";

        }

        // get iframe
        var iframe = document.getElementById('aemPage');
        console.log(iframe);
        // reload iframe with preview url
        console.log(json);
        iframe.src = `${json.live.url}?date=${Date.now()}`;
        iframe.addEventListener('load', handleLoad, true)
      });



    await context.sync();

  });
}


// function for checkConfig
export async function checkConfig() {
  return Word.run(async (context) => {
    var aemRepo = Office.context.document.settings.get('aemRepo');
    var productionUrl = Office.context.document.settings.get('productionUrl');
    var contentUrl = Office.context.document.settings.get('contentUrl');

    var pageOptions = document.getElementById('pageOptions');
    var config = document.getElementById('config');
    var iframe = document.getElementById('aemPage');
    var header = document.getElementById('aemHeader');
    console.log('config checked...');

    console.log(aemRepo);
    console.log(productionUrl);
    console.log(contentUrl);


    if (aemRepo && productionUrl && contentUrl) {
      getInitialState();

      pageOptions.classList.remove('d-none');
      config.classList.add('d-none');
      header.classList.add('d-none');
      iframe.classList.remove('d-none');
    } else {
      pageOptions.classList.add('d-none');
      config.classList.remove('d-none');
      header.classList.remove('d-none');
      iframe.classList.add('d-none');
    }
  });
}

export async function saveConfig() {
  return Word.run(async (context) => {
    var aemRepo = document.getElementById('aemRepo').value;
    var productionUrl = document.getElementById('productionUrl').value;
    var contentUrl = document.getElementById('contentUrl').value;

    Office.context.document.settings.set('aemRepo', aemRepo);
    Office.context.document.settings.set('productionUrl', productionUrl);
    Office.context.document.settings.set('contentUrl', contentUrl);

    Office.context.document.settings.saveAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log('Settings save failed. Error: ' + asyncResult.error.message);
      } else {
        console.log('Settings saved.');
      }
    });

    checkConfig();
    await context.sync();
  });
}


export async function editConfig() {
  return Word.run(async (context) => {

    var pageOptions = document.getElementById('pageOptions');
    var config = document.getElementById('config');
    var iframe = document.getElementById('aemPage');
    var header = document.getElementById('aemHeader');
    var pageMetadata = document.getElementById('pageMetadata');

    pageMetadata.classList.add('d-none');
    header.classList.remove('d-none');
    iframe.classList.add('d-none');
    pageOptions.classList.add('d-none');
    config.classList.remove('d-none');
    await context.sync();
  });
}

