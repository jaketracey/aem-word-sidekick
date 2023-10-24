/*! For license information please see taskpane.js.LICENSE.txt */
!function(){var t={27091:function(t){"use strict";t.exports=function(t,e){return e||(e={}),t?(t=String(t.__esModule?t.default:t),e.hash&&(t+=e.hash),e.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(t)?'"'.concat(t,'"'):t):t}},35666:function(t){var e=function(t){"use strict";var e,n=Object.prototype,r=n.hasOwnProperty,o=Object.defineProperty||function(t,e,n){t[e]=n.value},i="function"==typeof Symbol?Symbol:{},c=i.iterator||"@@iterator",a=i.asyncIterator||"@@asyncIterator",u=i.toStringTag||"@@toStringTag";function s(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{s({},"")}catch(t){s=function(t,e,n){return t[e]=n}}function l(t,e,n,r){var i=e&&e.prototype instanceof v?e:v,c=Object.create(i.prototype),a=new P(r||[]);return o(c,"_invoke",{value:_(t,n,a)}),c}function f(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}t.wrap=l;var p="suspendedStart",h="suspendedYield",d="executing",y="completed",m={};function v(){}function g(){}function w(){}var x={};s(x,c,(function(){return this}));var b=Object.getPrototypeOf,L=b&&b(b(S([])));L&&L!==n&&r.call(L,c)&&(x=L);var E=w.prototype=v.prototype=Object.create(x);function O(t){["next","throw","return"].forEach((function(e){s(t,e,(function(t){return this._invoke(e,t)}))}))}function I(t,e){function n(o,i,c,a){var u=f(t[o],t,i);if("throw"!==u.type){var s=u.arg,l=s.value;return l&&"object"==typeof l&&r.call(l,"__await")?e.resolve(l.__await).then((function(t){n("next",t,c,a)}),(function(t){n("throw",t,c,a)})):e.resolve(l).then((function(t){s.value=t,c(s)}),(function(t){return n("throw",t,c,a)}))}a(u.arg)}var i;o(this,"_invoke",{value:function(t,r){function o(){return new e((function(e,o){n(t,r,e,o)}))}return i=i?i.then(o,o):o()}})}function _(t,e,n){var r=p;return function(o,i){if(r===d)throw new Error("Generator is already running");if(r===y){if("throw"===o)throw i;return T()}for(n.method=o,n.arg=i;;){var c=n.delegate;if(c){var a=k(c,n);if(a){if(a===m)continue;return a}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(r===p)throw r=y,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);r=d;var u=f(t,e,n);if("normal"===u.type){if(r=n.done?y:h,u.arg===m)continue;return{value:u.arg,done:n.done}}"throw"===u.type&&(r=y,n.method="throw",n.arg=u.arg)}}}function k(t,n){var r=n.method,o=t.iterator[r];if(o===e)return n.delegate=null,"throw"===r&&t.iterator.return&&(n.method="return",n.arg=e,k(t,n),"throw"===n.method)||"return"!==r&&(n.method="throw",n.arg=new TypeError("The iterator does not provide a '"+r+"' method")),m;var i=f(o,t.iterator,n.arg);if("throw"===i.type)return n.method="throw",n.arg=i.arg,n.delegate=null,m;var c=i.arg;return c?c.done?(n[t.resultName]=c.value,n.next=t.nextLoc,"return"!==n.method&&(n.method="next",n.arg=e),n.delegate=null,m):c:(n.method="throw",n.arg=new TypeError("iterator result is not an object"),n.delegate=null,m)}function j(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function B(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function P(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(j,this),this.reset(!0)}function S(t){if(t){var n=t[c];if(n)return n.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var o=-1,i=function n(){for(;++o<t.length;)if(r.call(t,o))return n.value=t[o],n.done=!1,n;return n.value=e,n.done=!0,n};return i.next=i}}return{next:T}}function T(){return{value:e,done:!0}}return g.prototype=w,o(E,"constructor",{value:w,configurable:!0}),o(w,"constructor",{value:g,configurable:!0}),g.displayName=s(w,u,"GeneratorFunction"),t.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===g||"GeneratorFunction"===(e.displayName||e.name))},t.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,w):(t.__proto__=w,s(t,u,"GeneratorFunction")),t.prototype=Object.create(E),t},t.awrap=function(t){return{__await:t}},O(I.prototype),s(I.prototype,a,(function(){return this})),t.AsyncIterator=I,t.async=function(e,n,r,o,i){void 0===i&&(i=Promise);var c=new I(l(e,n,r,o),i);return t.isGeneratorFunction(n)?c:c.next().then((function(t){return t.done?t.value:c.next()}))},O(E),s(E,u,"Generator"),s(E,c,(function(){return this})),s(E,"toString",(function(){return"[object Generator]"})),t.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},t.values=S,P.prototype={constructor:P,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=e,this.done=!1,this.delegate=null,this.method="next",this.arg=e,this.tryEntries.forEach(B),!t)for(var n in this)"t"===n.charAt(0)&&r.call(this,n)&&!isNaN(+n.slice(1))&&(this[n]=e)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var n=this;function o(r,o){return a.type="throw",a.arg=t,n.next=r,o&&(n.method="next",n.arg=e),!!o}for(var i=this.tryEntries.length-1;i>=0;--i){var c=this.tryEntries[i],a=c.completion;if("root"===c.tryLoc)return o("end");if(c.tryLoc<=this.prev){var u=r.call(c,"catchLoc"),s=r.call(c,"finallyLoc");if(u&&s){if(this.prev<c.catchLoc)return o(c.catchLoc,!0);if(this.prev<c.finallyLoc)return o(c.finallyLoc)}else if(u){if(this.prev<c.catchLoc)return o(c.catchLoc,!0)}else{if(!s)throw new Error("try statement without catch or finally");if(this.prev<c.finallyLoc)return o(c.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var o=this.tryEntries[n];if(o.tryLoc<=this.prev&&r.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var i=o;break}}i&&("break"===t||"continue"===t)&&i.tryLoc<=e&&e<=i.finallyLoc&&(i=null);var c=i?i.completion:{};return c.type=t,c.arg=e,i?(this.method="next",this.next=i.finallyLoc,m):this.complete(c)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),m},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),B(n),m}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;B(n)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,n,r){return this.delegate={iterator:S(t),resultName:n,nextLoc:r},"next"===this.method&&(this.arg=e),m}},t}(t.exports);try{regeneratorRuntime=e}catch(t){"object"==typeof globalThis?globalThis.regeneratorRuntime=e:Function("r","regeneratorRuntime = r")(e)}},44944:function(t,e,n){"use strict";t.exports=n.p+"assets/logo-filled.png"},60806:function(t,e,n){"use strict";t.exports=n.p+"a47cb16a3177fc8ac49b.css"}},e={};function n(r){var o=e[r];if(void 0!==o)return o.exports;var i=e[r]={exports:{}};return t[r](i,i.exports,n),i.exports}n.m=t,n.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return n.d(e,{a:e}),e},n.d=function(t,e){for(var r in e)n.o(e,r)&&!n.o(t,r)&&Object.defineProperty(t,r,{enumerable:!0,get:e[r]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),n.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;n.g.importScripts&&(t=n.g.location+"");var e=n.g.document;if(!t&&e&&(e.currentScript&&(t=e.currentScript.src),!t)){var r=e.getElementsByTagName("script");if(r.length)for(var o=r.length-1;o>-1&&!t;)t=r[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=t}(),n.b=document.baseURI||self.location.href,function(){"use strict";function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(){e=function(){return r};var n,r={},o=Object.prototype,i=o.hasOwnProperty,c=Object.defineProperty||function(t,e,n){t[e]=n.value},a="function"==typeof Symbol?Symbol:{},u=a.iterator||"@@iterator",s=a.asyncIterator||"@@asyncIterator",l=a.toStringTag||"@@toStringTag";function f(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{f({},"")}catch(n){f=function(t,e,n){return t[e]=n}}function p(t,e,n,r){var o=e&&e.prototype instanceof w?e:w,i=Object.create(o.prototype),a=new T(r||[]);return c(i,"_invoke",{value:j(t,n,a)}),i}function h(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}r.wrap=p;var d="suspendedStart",y="suspendedYield",m="executing",v="completed",g={};function w(){}function x(){}function b(){}var L={};f(L,u,(function(){return this}));var E=Object.getPrototypeOf,O=E&&E(E(N([])));O&&O!==o&&i.call(O,u)&&(L=O);var I=b.prototype=w.prototype=Object.create(L);function _(t){["next","throw","return"].forEach((function(e){f(t,e,(function(t){return this._invoke(e,t)}))}))}function k(e,n){function r(o,c,a,u){var s=h(e[o],e,c);if("throw"!==s.type){var l=s.arg,f=l.value;return f&&"object"==t(f)&&i.call(f,"__await")?n.resolve(f.__await).then((function(t){r("next",t,a,u)}),(function(t){r("throw",t,a,u)})):n.resolve(f).then((function(t){l.value=t,a(l)}),(function(t){return r("throw",t,a,u)}))}u(s.arg)}var o;c(this,"_invoke",{value:function(t,e){function i(){return new n((function(n,o){r(t,e,n,o)}))}return o=o?o.then(i,i):i()}})}function j(t,e,r){var o=d;return function(i,c){if(o===m)throw new Error("Generator is already running");if(o===v){if("throw"===i)throw c;return{value:n,done:!0}}for(r.method=i,r.arg=c;;){var a=r.delegate;if(a){var u=B(a,r);if(u){if(u===g)continue;return u}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===d)throw o=v,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=m;var s=h(t,e,r);if("normal"===s.type){if(o=r.done?v:y,s.arg===g)continue;return{value:s.arg,done:r.done}}"throw"===s.type&&(o=v,r.method="throw",r.arg=s.arg)}}}function B(t,e){var r=e.method,o=t.iterator[r];if(o===n)return e.delegate=null,"throw"===r&&t.iterator.return&&(e.method="return",e.arg=n,B(t,e),"throw"===e.method)||"return"!==r&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+r+"' method")),g;var i=h(o,t.iterator,e.arg);if("throw"===i.type)return e.method="throw",e.arg=i.arg,e.delegate=null,g;var c=i.arg;return c?c.done?(e[t.resultName]=c.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=n),e.delegate=null,g):c:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,g)}function P(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function S(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function T(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(P,this),this.reset(!0)}function N(e){if(e||""===e){var r=e[u];if(r)return r.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,c=function t(){for(;++o<e.length;)if(i.call(e,o))return t.value=e[o],t.done=!1,t;return t.value=n,t.done=!0,t};return c.next=c}}throw new TypeError(t(e)+" is not iterable")}return x.prototype=b,c(I,"constructor",{value:b,configurable:!0}),c(b,"constructor",{value:x,configurable:!0}),x.displayName=f(b,l,"GeneratorFunction"),r.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===x||"GeneratorFunction"===(e.displayName||e.name))},r.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,b):(t.__proto__=b,f(t,l,"GeneratorFunction")),t.prototype=Object.create(I),t},r.awrap=function(t){return{__await:t}},_(k.prototype),f(k.prototype,s,(function(){return this})),r.AsyncIterator=k,r.async=function(t,e,n,o,i){void 0===i&&(i=Promise);var c=new k(p(t,e,n,o),i);return r.isGeneratorFunction(e)?c:c.next().then((function(t){return t.done?t.value:c.next()}))},_(I),f(I,l,"Generator"),f(I,u,(function(){return this})),f(I,"toString",(function(){return"[object Generator]"})),r.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},r.values=N,T.prototype={constructor:T,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=n,this.done=!1,this.delegate=null,this.method="next",this.arg=n,this.tryEntries.forEach(S),!t)for(var e in this)"t"===e.charAt(0)&&i.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=n)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function r(r,o){return a.type="throw",a.arg=t,e.next=r,o&&(e.method="next",e.arg=n),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var c=this.tryEntries[o],a=c.completion;if("root"===c.tryLoc)return r("end");if(c.tryLoc<=this.prev){var u=i.call(c,"catchLoc"),s=i.call(c,"finallyLoc");if(u&&s){if(this.prev<c.catchLoc)return r(c.catchLoc,!0);if(this.prev<c.finallyLoc)return r(c.finallyLoc)}else if(u){if(this.prev<c.catchLoc)return r(c.catchLoc,!0)}else{if(!s)throw new Error("try statement without catch or finally");if(this.prev<c.finallyLoc)return r(c.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var r=this.tryEntries[n];if(r.tryLoc<=this.prev&&i.call(r,"finallyLoc")&&this.prev<r.finallyLoc){var o=r;break}}o&&("break"===t||"continue"===t)&&o.tryLoc<=e&&e<=o.finallyLoc&&(o=null);var c=o?o.completion:{};return c.type=t,c.arg=e,o?(this.method="next",this.next=o.finallyLoc,g):this.complete(c)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),g},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),S(n),g}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;S(n)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,e,r){return this.delegate={iterator:N(t),resultName:e,nextLoc:r},"next"===this.method&&(this.arg=n),g}},r}function r(t,e,n,r,o,i,c){try{var a=t[i](c),u=a.value}catch(t){return void n(t)}a.done?e(u):Promise.resolve(u).then(r,o)}function o(t){return function(){var e=this,n=arguments;return new Promise((function(o,i){var c=t.apply(e,n);function a(t){r(c,o,i,a,u,"next",t)}function u(t){r(c,o,i,a,u,"throw",t)}a(void 0)}))}}function i(){return c.apply(this,arguments)}function c(){return c=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(n){var r,o,i,c,a,u,s,l,f;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:r=Office.context.document.url,o=Office.context.document.settings.get("aemRepo"),i=o.replace("https://github.com/",""),c=Office.context.document.settings.get("productionUrl"),a=Office.context.document.settings.get("contentUrl"),u=document.getElementById("preview"),s=document.getElementById("publish"),l=document.getElementById("pageMetadata"),r=Office.context.document.url,f=document.getElementById("viewProduction"),r=(r=(r=(r=(r=(r=r.replace(" ","%20")).replace(a,"")).replace(/ /g,"-")).replace(".docx","")).replace(/’/g,"-")).toLowerCase(),fetch("https://admin.hlx.page/status/"+i+r,{method:"GET"}).then((function(t){return t.json()})).then((function(t){document.getElementById("lastModified").innerHTML="Last modified: ".concat(t.preview.lastModified);var e=document.getElementById("aemPage");e.src="".concat(t.preview.url,"?date=").concat(Date.now()),e.addEventListener("load",(function(){loader.classList.add("d-none"),u.textContent="Preview",s.textContent="Publish",l.classList.remove("d-none"),document.getElementById("pageOptions").classList.remove("d-none")}),!0),t.live.url&&(f.classList.remove("d-none"),f.addEventListener("click",(function(){if(c){var e=new URL(t.live.url);window.open("https://".concat(c+e.pathname),"_blank")}else window.open(t.live.url,"_blank")})))}));case 18:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)}))),c.apply(this,arguments)}function a(){return u.apply(this,arguments)}function u(){return u=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(n){var r,o,i,c,a,u;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return r=Office.context.document.url,o=Office.context.document.settings.get("aemRepo"),i=o.replace("https://github.com/",""),Office.context.document.settings.get("productionUrl"),c=Office.context.document.settings.get("contentUrl"),(a=document.getElementById("preview")).textContent="Previewing...",(u=document.getElementById("loader")).classList.remove("d-none"),r=(r=(r=(r=(r=(r=r.replace(" ","%20")).replace(c,"")).replace(/ /g,"-")).replace(".docx","")).replace(/’/g,"-")).toLowerCase(),fetch("https://admin.hlx.page/preview/"+i+r,{method:"POST"}).then((function(t){return t.json()})).then((function(t){document.getElementById("lastModified").innerHTML="Last modified: ".concat(t.preview.lastModified);var e=document.getElementById("aemPage");e.src="".concat(t.preview.url,"?date=").concat(Date.now()),e.addEventListener("load",(function(){u.classList.add("d-none"),a.textContent="Preview"}),!0)})),t.next=19,n.sync();case 19:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)}))),u.apply(this,arguments)}function s(){return l.apply(this,arguments)}function l(){return l=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(n){var r,o,i,c,a,u;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return r=Office.context.document.url,o=Office.context.document.settings.get("aemRepo"),i=o.replace("https://github.com/",""),Office.context.document.settings.get("productionUrl"),c=Office.context.document.settings.get("contentUrl"),(a=document.getElementById("publish")).textContent="Publishing...",(u=document.getElementById("loader")).classList.remove("d-none"),r=(r=(r=(r=(r=(r=r.replace(" ","%20")).replace(c,"")).replace(/ /g,"-")).replace(".docx","")).replace(/’/g,"-")).toLowerCase(),fetch("https://admin.hlx.page/live/"+i+r,{method:"POST",body:null}).then((function(t){return t.json()})).then((function(t){document.getElementById("lastModified").innerHTML="Last modified: ".concat(t.live.lastModified);var e=document.getElementById("aemPage");e.src="".concat(t.live.url,"?date=").concat(Date.now()),e.addEventListener("load",(function(){u.classList.add("d-none"),a.textContent="Publish"}),!0)})),t.next=19,n.sync();case 19:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)}))),l.apply(this,arguments)}function f(){return p.apply(this,arguments)}function p(){return p=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(n){var r,o,c,a,u;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:r=Office.context.document.settings.get("aemRepo"),Office.context.document.settings.get("productionUrl"),o=Office.context.document.settings.get("contentUrl"),c=document.getElementById("config"),a=document.getElementById("aemPage"),u=document.getElementById("aemHeader"),r&&o?(i(),c.classList.add("d-none"),u.classList.add("d-none"),a.classList.remove("d-none")):(c.classList.remove("d-none"),u.classList.remove("d-none"),a.classList.add("d-none"));case 7:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)}))),p.apply(this,arguments)}function h(){return d.apply(this,arguments)}function d(){return d=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(n){var r,o,i;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return r=document.getElementById("aemRepo").value,o=document.getElementById("productionUrl").value,i=document.getElementById("contentUrl").value,Office.context.document.settings.set("aemRepo",r),Office.context.document.settings.set("productionUrl",o),Office.context.document.settings.set("contentUrl",i),Office.context.document.settings.saveAsync((function(t){t.status==Office.AsyncResultStatus.Failed?console.log("Settings save failed. Error: "+t.error.message):console.log("Settings saved.")})),f(),t.next=10,n.sync();case 10:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)}))),d.apply(this,arguments)}function y(){return m.apply(this,arguments)}function m(){return m=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(n){var r,o,i,c;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return r=document.getElementById("pageOptions"),o=document.getElementById("config"),i=document.getElementById("aemPage"),c=document.getElementById("aemHeader"),document.getElementById("pageMetadata").classList.add("d-none"),c.classList.remove("d-none"),i.classList.add("d-none"),r.classList.add("d-none"),o.classList.remove("d-none"),t.next=12,n.sync();case 12:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)}))),m.apply(this,arguments)}n(35666),Office.onReady((function(t){t.host===Office.HostType.Word&&(document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("editConfig").onclick=y,document.getElementById("saveConfig").onclick=h,document.getElementById("publish").onclick=s,document.getElementById("preview").onclick=a),f()}))}(),function(){"use strict";var t=n(27091),e=n.n(t),r=new URL(n(60806),n.b),o=new URL(n(44944),n.b);e()(r),e()(o)}()}();
//# sourceMappingURL=taskpane.js.map