import { async } from "regenerator-runtime";

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";

  }

  /*
    Define buttons to be in the page options.
    Each button has a label and an action.
    The action is the name of a function that will be called when the button is clicked.
    If the button has an icon the label will not be displayed.
  */
  var buttons = {
    'preview': {
      'label': 'Preview',
      'action': 'preview'
    },
    'publish': {
      'label': 'Publish',
      'id': 'publish'
    },
    // 'unpublish': {
    //   'label': 'Unpublish',
    //   'id': 'unpublish',
    // },
    'viewProduction': {
      'label': 'View Production',
      'id': 'viewProduction',
      'icon': 'ms-Icon--OpenInNewWindow'
    },
    'viewLibrary': {
      'label': 'View Library',
      'id': 'viewLibrary',
      'icon': 'ms-Icon--Library'
    },
    'editConfig': {
      'label': 'Edit Config',
      'id': 'editConfig',
      'icon': 'ms-Icon--Settings'
    },
  };

  var aemRepo;
  var aemRepoName;
  var fileUrl;
  var liveUrl;
  var previewUrl;
  var firstRun = document.getElementById('first-run');
  var contentUrl = Office.context.document.settings.get('contentUrl');
  var config = document.getElementById('config');
  var iframe = document.getElementById('aemPage');
  var pageMetadata = document.getElementById('pageMetadata');
  var pageOptions = document.getElementById('pageOptions');
  var productionUrl = Office.context.document.settings.get('productionUrl');
  var loader = document.getElementById("loader");
  var smallLoader = document.getElementById("small-loader");


  // find element with id saveConfig
  var saveConfig = document.getElementById('saveConfig');
  // add click event to the saveConfig button
  saveConfig.addEventListener('click', function () {
    actions.saveConfig();
  });


  var lastPublished = document.getElementById('lastPublished');
  var lastPreviewed = document.getElementById('lastPreviewed');
  var lastModified = document.getElementById('lastModified');

  pageMetadata.addEventListener('click', function (e) {
    e.stopPropagation();
    pageMetadata.classList.toggle('expanded');
  });

  // create buttons for each button in buttons and attach an onclick event
  for (var key in buttons) {
    var button = document.createElement('button');
    button.classList.add('ms-Button');
    button.setAttribute('id', key);
    button.setAttribute('type', 'button');
    button.setAttribute('name', key);

    // if the button has an icon add it
    if (buttons[key].icon) {
      button.classList.add('ms-Button-withIcon');
      button.innerHTML = `<span class="ms-Button-icon"><i class="ms-Icon ${buttons[key].icon}"></i></span><span class="ms-Button-label">${buttons[key].label}</span>`;
      document.getElementById('pageOptions-icons').appendChild(button);
    } else {
      button.innerHTML = `<span class="ms-Button-label">${buttons[key].label}</span>`;
      document.getElementById('pageOptions-actions').appendChild(button);
    }

    // add click event to the button
    button.addEventListener('click', function (e) {
      e.stopPropagation();
      // get the name of the button that was clicked
      var action = e.currentTarget.getAttribute('id');

      // call a function with the name of the button that was clicked
      actions[`${action}`]();
    });

    ///
    // All action functions live here
    //
    var actions = {
      publish: async function () {
        return Word.run(async (context) => {
          var publishButton = document.getElementById("publish");
          publishButton.setAttribute('disabled', 'disabled');
          publishButton.textContent = "Publishing...";
          publishButton.classList.add('disabled');
          loader.classList.remove('d-none');

          var liveUrl = 'https://admin.hlx.page/live/' + aemRepoName + fileUrl;
          fetch(liveUrl, {
            method: "POST",
            body: null,
          })
            .then((response) => response.json())
            .then((json) => {

              // update page metadata
              updatePageMetadata();

              function handleLoad() {
                pageMetadata.classList.remove('d-none');
                pageOptions.classList.remove('d-none');
                iframe.classList.remove('d-none');
                smallLoader.classList.add('d-none');
                loader.classList.add('d-none');
                publishButton.textContent = "Publish";
                publishButton.removeAttribute('disabled');
                publishButton.classList.remove('disabled');
              }


              iframe.src = `${json.live.url}?date=${Date.now()}`;
              iframe.addEventListener('load', handleLoad, true)
            });

          await context.sync();
        });
      },
      preview: async function () {
        return Word.run(async (context) => {
          var previewButton = document.getElementById("preview");
          previewButton.textContent = "Previewing...";
          previewButton.setAttribute('disabled', 'disabled');
          previewButton.classList.add('disabled');

          loader.classList.remove('d-none');

          var liveUrl = 'https://admin.hlx.page/preview/' + aemRepoName + fileUrl;
          fetch(liveUrl, {
            method: "POST",
          })
            .then((response) => response.json())
            .then((json) => {

              // update page metadata
              updatePageMetadata();

              function handleLoad() {
                pageMetadata.classList.remove('d-none');
                pageOptions.classList.remove('d-none');
                iframe.classList.remove('d-none');
                smallLoader.classList.add('d-none');
                loader.classList.add('d-none');
                previewButton.textContent = "Preview";
                previewButton.removeAttribute('disabled');
                previewButton.classList.remove('disabled');
              }

              // reload iframe with preview url
              iframe.src = `${json.preview.url}?date=${Date.now()}`;
              iframe.addEventListener('load', handleLoad, true)
            });

          await context.sync();
        });
      },
      unpublish: async function () {
        return Word.run(async (context) => {

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
                if (json.live.lastModified) {
                  lastModified.innerHTML = `Last modified: ${json.live.lastModified}`;
                } else {
                  lastModified.innerHTML = `No page published yet`;
                }

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
      },
      viewProduction: async function () {
        // if productionUrl is set in the config use it
        if (productionUrl) {
          // strip the domain from json.live.url
          var url = new URL(liveUrl);

          window.open(`https://${productionUrl + url.pathname}`, '_blank');
        } else {
          // otherwise use the live url from the api
          window.open(liveUrl, '_blank');
        }
      },

      editConfig: async function () {
        return Word.run(async (context) => {

          firstRun.classList.add('d-none');

          // read the settings
          aemRepo = Office.context.document.settings.get('aemRepo');
          productionUrl = Office.context.document.settings.get('productionUrl');
          contentUrl = Office.context.document.settings.get('contentUrl');


          if (productionUrl) {
            document.getElementById('productionUrl').value = productionUrl;
          }
          document.getElementById('contentUrl').value = contentUrl;
          document.getElementById('aemRepo').value = aemRepo;

          loader.classList.add('d-none');
          smallLoader.classList.add('d-none');
          pageMetadata.classList.add('d-none');
          iframe.classList.add('d-none');
          pageOptions.classList.add('d-none');
          config.classList.remove('d-none');
          configRibbon = false;

          await context.sync();
        });
      },
      saveConfig: async function () {
        return Word.run(async (context) => {
          var configError = document.getElementById('config-error');
          configError.classList.add('d-none');

          if (aemRepo == '' || contentUrl == '') {
            configError.innerHTML = 'Please enter both Github repo and Content URL fields';
            configError.classList.remove('d-none');
          }

          // get the values from the config form
          aemRepo = document.getElementById('aemRepo').value;
          productionUrl = document.getElementById('productionUrl').value;
          contentUrl = document.getElementById('contentUrl').value;

          // we need to check if the filename has changed, so we have to get the current fileUrl values again
          fileUrl = getFormattedDocumentUrl();
          liveUrl = 'https://admin.hlx.page/live/' + aemRepoName + fileUrl;

          Office.context.document.settings.set('aemRepo', aemRepo);
          Office.context.document.settings.set('productionUrl', productionUrl);
          Office.context.document.settings.set('contentUrl', contentUrl);

          Office.context.document.settings.saveAsync(function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
              console.log('Settings save failed. Error: ' + asyncResult.error.message);
            } else {
              console.log('Settings saved.');
              console.log(`aemRepo: ${aemRepo}`);
              console.log(`productionUrl: ${productionUrl}`);
              console.log(`contentUrl: ${contentUrl}`);
            }
          });

          actions.checkConfig();
          await context.sync();
        });
      },

      checkConfig: async function () {
        return Word.run(async (context) => {
          console.log('checking if config exists');

          // if url doesn't contain https:// then it's a local file
          if (Office.context.document.url.indexOf('https://') == -1) {

            // show the first run screen
            firstRun.classList.remove('d-none');
            firstRun.innerHTML = `<h3>Local file detected</h3>
            <p>This add in only supports documents located in Sharepoint.</p><p>Please open your document from a SharePoint location to continue.</p>`;
            // find element of config-body id
            var configBody = document.getElementById('config-body');
            // hide the config
            configBody.classList.add('d-none');

            pageMetadata.classList.add('d-none');
            pageOptions.classList.add('d-none');
            iframe.classList.add('d-none');
            loader.classList.add('d-none');
            config.classList.remove('d-none');
          }

          aemRepo = Office.context.document.settings.get('aemRepo');
          productionUrl = Office.context.document.settings.get('productionUrl');
          contentUrl = Office.context.document.settings.get('contentUrl');

          console.log(`aemRepo: ${aemRepo}`);
          console.log(`productionUrl: ${productionUrl}`);
          console.log(`contentUrl: ${contentUrl}`);

          if (aemRepo && contentUrl) {

              actions.getInitialState(aemRepo);
              config.classList.add('d-none');
            iframe.classList.remove('d-none');
            loader.classList.remove('d-none');
          } else {
            config.classList.remove('d-none');
            iframe.classList.add('d-none');
          }
        });
      },
      viewLibrary: async function () {
        window.open(`https://${previewUrl.hostname}/tools/sidekick/library.html`, '_blank');
      },
      getInitialState: async function () {
        return Word.run(async (context) => {
          fileUrl = getFormattedDocumentUrl();
          aemRepo = Office.context.document.settings.get('aemRepo');
          aemRepoName = aemRepo.replace('https://github.com/', '');
          liveUrl = 'https://admin.hlx.page/live/' + aemRepoName + fileUrl;
          productionUrl = Office.context.document.settings.get('productionUrl');
          contentUrl = Office.context.document.settings.get('productionUrl');



          var previewButton = document.getElementById("preview");
          var publishButton = document.getElementById("publish");
          var pageMetadata = document.getElementById("pageMetadata");
          var viewProductionButton = document.getElementById("viewProduction");
          var pageOptions = document.getElementById('pageOptions');
          var viewLibrary = document.getElementById('viewLibrary');


          var iframe = document.getElementById('aemPage');
          iframe.classList.add('d-none');


          console.log(contentUrl);
          console.log(fileUrl);




          // if publishRibbon or previewRibbon is set to true then show the ribbon
            // if preview is set then preview the page
            if (previewRibbon) {
              actions.preview();
              smallLoader.classList.add('d-none');
              loader.classList.add('d-none');
            } else if(publishRibbon) {
              actions.publish();
              smallLoader.classList.add('d-none');
              loader.classList.add('d-none');
            }else if (configRibbon) {
              console.log('yeet');
            actions.editConfig();
          } else {
          var statusEndpoint = 'https://admin.hlx.page/status/' + aemRepoName + fileUrl;

          fetch(statusEndpoint, {
            method: "GET",
          })
            .then((response) => {

              if (response.status == 404) {
                // show the first run screen
                firstRun.classList.remove('d-none');
                firstRun.innerHTML = `<h3>No page found</h3><p>We can't find the page on AEM you're editing.</p><p>Please preview the page to use this add in.</p><p>You can edit files that are assigned to your AEM project from Sharepoint.</p>`;
                pageMetadata.classList.add('d-none');
                pageOptions.classList.add('d-none');
                iframe.classList.add('d-none');
                loader.classList.add('d-none');
                config.classList.remove('d-none');

                return;
              } else {
                return response.json()

              }
            })
            .then((json) => {
              console.log(json);
              // get iframe
              // reload iframe with preview url
              iframe.src = `${json.preview.url}?date=${Date.now()}`;
              iframe.addEventListener('load', handleLoad, true);

              // update page metadata
              updatePageMetadata(json);

              // show the view button if the page is published
              if (json.live.url) {
                viewProductionButton.classList.remove('d-none');
                liveUrl = json.live.url;

              }

              previewUrl = new URL(json.preview.url);


              function handleLoad() {

                iframe.classList.remove('d-none');
                loader.classList.add('d-none');
                previewButton.textContent = "Preview";
                publishButton.textContent = "Publish";
                pageMetadata.classList.remove('d-none');
                pageOptions.classList.remove('d-none');
              }
            });
          }
        });
      },

    };

  }

  /*
    Update the page metadata based on the json response
  */
  function updatePageMetadata() {

    var statusEndpoint = 'https://admin.hlx.page/status/' + aemRepoName + fileUrl;

    fetch(statusEndpoint, {
      method: "GET",
    })
      .then((response) => response.json())
      .then((json) => {
        // convert date to local time
        var lastModifiedString = new Date(json.live.sourceLastModified).toLocaleString("en-AU", {
          day: 'numeric',
          month: 'long',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit'

        });;
        var lastPublishedString = new Date(json.live.lastModified).toLocaleString("en-AU", {
          day: 'numeric',
          month: 'long',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit'
        });;
        var lastPreviewedString = new Date(json.preview.lastModified).toLocaleString("en-AU", {
          day: 'numeric',
          month: 'long',
          year: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit'
        });;

        lastModified.innerHTML = `Modified: ${lastModifiedString}`;

        lastPublished.innerHTML = `Published: ${lastPublishedString}`;

        lastPreviewed.innerHTML = `Previewed: ${lastPreviewedString}`;

        // remove advancedOptions if it exists
        var advancedOptions = document.getElementById('advancedOptions');
        if (advancedOptions) {
          advancedOptions.remove();
        }

        // create advanced div
        var advanced = document.createElement('div');
        advanced.classList.add('advanced');
        advanced.id = 'advancedOptions';


        //pageMetadata.appendChild(advanced);

        // create clear cache button
        var clearCache = document.createElement('button');
        clearCache.classList.add('ms-Button');
        clearCache.setAttribute('id', 'clearCache');
        clearCache.setAttribute('type', 'button');
        clearCache.setAttribute('name', 'clearCache');
        clearCache.innerHTML = `<span class="ms-Button-label">Clear Cache</span>`;

        // add event for clear cache
        clearCache.addEventListener('click', function () {
          // send clear cache request to hlx.page
          var liveUrl = 'https://admin.hlx.page/clear/' + aemRepoName + fileUrl;
          fetch(liveUrl, {
            method: "POST",
          })
            .then((response) => response.json())
            .then((json) => {
              console.log(json);
            })
        });

        advanced.appendChild(clearCache);

        // create reindex button
        var reindex = document.createElement('button');
        reindex.classList.add('ms-Button');
        reindex.setAttribute('id', 'reindex');
        reindex.setAttribute('type', 'button');
        reindex.setAttribute('name', 'reindex');

        reindex.innerHTML = `<span class="ms-Button-label">Reindex</span>`;
        // add event for reindex
        reindex.addEventListener('click', function () {
          // send reindex request to hlx.page
          var liveUrl = 'https://admin.hlx.page/index/' + aemRepoName + fileUrl;
          fetch(liveUrl, {
            method: "POST",
          })
            .then((response) => response.json())
            .then((json) => {
              console.log(json);
            })
        });

        advanced.appendChild(reindex);

        // create deindex button
        var deindex = document.createElement('button');
        deindex.classList.add('ms-Button');
        deindex.setAttribute('id', 'deindex');
        deindex.setAttribute('type', 'button');
        deindex.setAttribute('name', 'deindex');

        deindex.innerHTML = `<span class="ms-Button-label">Deindex</span>`;
        // add event for deindex
        deindex.addEventListener('click', function () {
          // send deindex request to hlx.page
          var liveUrl = 'https://admin.hlx.page/index/' + aemRepoName + fileUrl;
          fetch(liveUrl, {
            method: "DELETE",
          })
            .then((response) => response.json())
            .then((json) => {
              console.log(json);
            })
        });

        advanced.appendChild(deindex);
      });
  }

  /*
    Helper function to get the actual document URL
  */
  function getFormattedDocumentUrl() {
    // get the fileUrl from the document
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

    return fileUrl;
  }


            // check if we're being called by the buttons in the ribbon
            var queryString = window.location.search;
            var urlParams = new URLSearchParams(queryString);
            var previewRibbon = urlParams.get('preview');
            var publishRibbon = urlParams.get('publish');
            var configRibbon = urlParams.get('config');

      actions.checkConfig();


});
