/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { async } from "regenerator-runtime";

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";

    document.getElementById("editConfig").onclick = editConfig;
    document.getElementById("saveConfig").onclick = saveConfig;
    document.getElementById("publish").onclick = publish;
    // document.getElementById("unpublish").onclick = unpublish;
    document.getElementById("preview").onclick = preview;
    document.getElementById("viewLibrary").onclick = viewLibrary;
  }

  checkConfig();
});


export async function run() {
  return Word.run(async (context) => {
    checkConfig();

    await context.sync();
  });
}

function getFormattedDocumentUrl() {
  var fileUrl = Office.context.document.url;
  var contentUrl = Office.context.document.settings.get('contentUrl');

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

  return fileUrl;
}

export async function getInitialState(aemRepo) {
  return Word.run(async (context) => {
    var fileUrl = getFormattedDocumentUrl();
    var aemRepoName = aemRepo.replace('https://github.com/', '');

    var productionUrl = Office.context.document.settings.get('productionUrl');
    var previewButton = document.getElementById("preview");
    var publishButton = document.getElementById("publish");
    var pageMetadata = document.getElementById("pageMetadata");
    var viewProductionButton = document.getElementById("viewProduction");
    var pageOptions = document.getElementById('pageOptions');
    var viewLibrary = document.getElementById('viewLibrary');


    var liveUrl = 'https://admin.hlx.page/status/' + aemRepoName + fileUrl;
    var iframe = document.getElementById('aemPage');
    iframe.classList.add('d-none');

    // create loader
    var loader = document.createElement('div');
    loader.classList.add('small-loader');
    loader.setAttribute('id', 'loader');
    loader.innerHTML = `<p>Loading project configuration...</p>`;

    // add loader to body
    document.body.appendChild(loader);

    fetch(liveUrl, {
      method: "GET",
    })
      .then((response) => response.json())
      .then((json) => {
        console.log(json);
        // find element with id lastModified
        var lastModified = document.getElementById('lastModified');

        // create span for last edited
        var lastEdited = document.createElement('span');
        lastEdited.innerHTML = `Last edited: ${json.preview.sourceLastModified}`;

        // create span for last published
        var lastPublished = document.createElement('span');
        lastPublished.innerHTML = `Last published: ${json.live.lastModified}`;

        // create span for last previewed
        var lastPreviewed = document.createElement('span');
        lastPreviewed.innerHTML = `Last previewed: ${json.preview.lastModified}`;

        // clear the pageMetadata
        pageMetadata.innerHTML = '';

        if (json.preview.lastModified && json.live.lastModified) {

          //pageMetadata.appendChild(lastEdited);
          pageMetadata.appendChild(lastPublished);
          //pageMetadata.appendChild(lastPreviewed);
        } else {

          pageMetadata.innerHTML = `No page published`;
        }
        // get iframe
        // reload iframe with preview url
        iframe.src = `${json.preview.url}?date=${Date.now()}`;
        iframe.addEventListener('load', handleLoad, true)
        // show the view button if the page is published
        if (json.live.url) {
          viewProductionButton.classList.remove('d-none');

          // add click event to the view button to open the page in a new tab
          viewProductionButton.addEventListener('click', function () {
            // if productionUrl is set in the config use it
            if (productionUrl) {
              // strip the domain from json.live.url
              var url = new URL(json.live.url);

              window.open(`https://${productionUrl + url.pathname}`, '_blank');
            } else {
              // otherwise use the live url from the api
              window.open(json.live.url, '_blank');
            }
          });

          // check if the library exists
          var libraryUrl = `https://${productionUrl}/tools/sidekick/library.html`;
          fetch(libraryUrl, {
            method: "GET",
          })
            .then((response) => {
              if (response.status == 200) {
                // add click event to the view button to open the page in a new tab
            viewLibrary.addEventListener('click', function () {
              // if productionUrl is set in the config use it

                // otherwise use the live url from the api
                window.open(`https://${url.hostname}/tools/sidekick/library.html`, '_blank');
            });

                viewLibrary.classList.remove('d-none');

              }
            })

        }

        function handleLoad() {

          iframe.classList.remove('d-none');
          loader.classList.add('d-none');
          previewButton.textContent = "Preview";
          publishButton.textContent = "Publish";
          pageMetadata.classList.remove('d-none');
          pageOptions.classList.remove('d-none');

        }
      });
  });
}



export async function preview() {
  return Word.run(async (context) => {
    var fileUrl = getFormattedDocumentUrl();
    var aemRepo = Office.context.document.settings.get('aemRepo');
    var aemRepoName = aemRepo.replace('https://github.com/', '');
    var previewButton = document.getElementById("preview");
    previewButton.textContent = "Previewing...";

    var loader = document.getElementById("loader");
    loader.classList.remove('d-none');

    var liveUrl = 'https://admin.hlx.page/preview/' + aemRepoName + fileUrl;
    fetch(liveUrl, {
      method: "POST",
    })
      .then((response) => response.json())
      .then((json) => {

        // clear pageMetadata
        var pageMetadata = document.getElementById('pageMetadata');
        pageMetadata.innerHTML = '';

        // create span for last edited
        var lastEdited = document.createElement('span');
        lastEdited.innerHTML = `Last edited: ${json.preview.sourceLastModified}`;

        pageMetadata.appendChild(lastEdited);

        function handleLoad() {
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
    var fileUrl = getFormattedDocumentUrl();
    var aemRepo = Office.context.document.settings.get('aemRepo');
    var aemRepoName = aemRepo.replace('https://github.com/', '');
    var productionUrl = Office.context.document.settings.get('productionUrl');
    var contentUrl = Office.context.document.settings.get('contentUrl');
    var publishButton = document.getElementById("publish");
    publishButton.textContent = "Publishing...";

    var loader = document.getElementById("loader");
    loader.classList.remove('d-none');

    var liveUrl = 'https://admin.hlx.page/live/' + aemRepoName + fileUrl;
    fetch(liveUrl, {
      method: "POST",
      body: null,
    })
      .then((response) => response.json())
      .then((json) => {

        // clear pageMetadata
        var pageMetadata = document.getElementById('pageMetadata');
        pageMetadata.innerHTML = '';

        console.log(json);

        // create span for last edited
        var lastEdited = document.createElement('span');
        lastEdited.innerHTML = `Last edited: ${json.live.sourceLastModified}`;

        pageMetadata.appendChild(lastEdited);

        function handleLoad() {
          loader.classList.add('d-none');
          publishButton.textContent = "Publish";

        }

        // get iframe
        var iframe = document.getElementById('aemPage');
        // reload iframe with preview url
        iframe.src = `${json.live.url}?date=${Date.now()}`;
        iframe.addEventListener('load', handleLoad, true)
      });

    await context.sync();

  });
}



export async function unpublish() {
  return Word.run(async (context) => {
    var fileUrl = getFormattedDocumentUrl();
    var aemRepo = Office.context.document.settings.get('aemRepo');
    var aemRepoName = aemRepo.replace('https://github.com/', '');

    // create element for a modal window and display it
    var modal = document.createElement('div');
    modal.classList.add('modal');
    modal.classList.add('fade');
    modal.classList.add('show');
    modal.setAttribute('id', 'unpublishModal');
    modal.setAttribute('tabindex', '-1');
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-labelledby', 'unpublishModalLabel');
    modal.setAttribute('aria-hidden', 'true');

    var modalContent = document.createElement('div');
    modalContent.classList.add('modal-content');

    var modalActions = document.createElement('div');
    modalActions.classList.add('modal-actions');

    modalContent.innerHTML = `<h2>Are you sure you want to unpublish this content?</h2>
    <p>Unpublishing content will make the page not visible for users</p>`;

    // create buttons for modal actions
    var unpublishConfirmButton = document.createElement('button');
    unpublishConfirmButton.classList.add('ms-Button-label');
    unpublishConfirmButton.setAttribute('id', 'unpublishConfirm');
    unpublishConfirmButton.textContent = 'Unpublish';



    var unpublishCancelButton = document.createElement('button');
    unpublishCancelButton.classList.add('ms-Button-label');
    unpublishCancelButton.setAttribute('id', 'unpublishCancel');
    unpublishCancelButton.setAttribute('data-dismiss', 'modal');
    unpublishCancelButton.textContent = 'Cancel';

    unpublishCancelButton.addEventListener('click', function () {
      // close the modal
      modal.classList.remove('show');
      modal.setAttribute('aria-hidden', 'true');
      modal.setAttribute('style', 'display: none');
      modal.setAttribute('aria-modal', 'false');
    });


    // create events for modal actions
    unpublishConfirmButton.addEventListener('click', function () {
      // close the modal
      modal.classList.remove('show');
      modal.setAttribute('aria-hidden', 'true');
      modal.setAttribute('style', 'display: none');
      modal.setAttribute('aria-modal', 'false');

      // sent unpublish request to hlx.page
      var liveUrl = 'https://admin.hlx.page/live/' + aemRepoName + fileUrl;
      fetch(liveUrl, {
        method: "DELETE",
      })
        .then((response) => {
          return response.json();
        })
        .then((json) => {
          // find element with id lastModified
          var lastModified = document.getElementById('lastModified');
          if (json.live.lastModified) {
            lastModified.innerHTML = `Last modified: ${json.live.lastModified}`;
          } else {
            lastModified.innerHTML = `No page published yet`;
          }

          // get iframe
          var iframe = document.getElementById('aemPage');
          // reload iframe with preview url
          iframe.src = `${json.live.url}?date=${Date.now()}`;
        })
    });


    // add buttons to modal actions
    modalActions.appendChild(unpublishConfirmButton);
    modalActions.appendChild(unpublishCancelButton);

    // add modal actions to modal content
    modalContent.appendChild(modalActions);

    // add modal content to modal
    modal.appendChild(modalContent);

    // add modal to body
    document.body.appendChild(modal);


    await context.sync();

  });
}


// function for checkConfig
export async function checkConfig() {
  return Word.run(async (context) => {
    var aemRepo = Office.context.document.settings.get('aemRepo');
    var contentUrl = Office.context.document.settings.get('contentUrl');

    var config = document.getElementById('config');
    var iframe = document.getElementById('aemPage');
    var loader = document.getElementById("loader");


    if (aemRepo && contentUrl) {
      getInitialState(aemRepo);

      config.classList.add('d-none');
      iframe.classList.remove('d-none');
    } else {
      config.classList.remove('d-none');
      iframe.classList.add('d-none');
    }
  });
}

export async function saveConfig() {
  return Word.run(async (context) => {
    var aemRepo = document.getElementById('aemRepo').value;
    var productionUrl = document.getElementById('productionUrl').value;
    var contentUrl = document.getElementById('contentUrl').value;
    var configError = document.getElementById('config-error');
    configError.classList.add('d-none');

    if (aemRepo == '' || contentUrl == '') {
      configError.innerHTML = 'Please enter both Github repo and Content URL fields';
      configError.classList.remove('d-none');
    }

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
    var pageMetadata = document.getElementById('pageMetadata');

    var firstRun = document.getElementById('first-run');
    firstRun.classList.add('d-none');

    // get the values from the settings
    var aemRepo = Office.context.document.settings.get('aemRepo');
    var productionUrl = Office.context.document.settings.get('productionUrl');
    var contentUrl = Office.context.document.settings.get('contentUrl');

    // populate the inputs with the values if they exist
    if (productionUrl) {
      document.getElementById('productionUrl').value = productionUrl;
    }
    document.getElementById('contentUrl').value = contentUrl;
    document.getElementById('aemRepo').value = aemRepo;


    pageMetadata.classList.add('d-none');
    iframe.classList.add('d-none');
    pageOptions.classList.add('d-none');
    config.classList.remove('d-none');

    await context.sync();
  });
}

