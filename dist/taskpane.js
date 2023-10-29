/*! For license information please see taskpane.js.LICENSE.txt */
!function(){var t={27091:function(t){"use strict";t.exports=function(t,e){return e||(e={}),t?(t=String(t.__esModule?t.default:t),e.hash&&(t+=e.hash),e.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(t)?'"'.concat(t,'"'):t):t}},35666:function(t){var e=function(t){"use strict";var e,n=Object.prototype,r=n.hasOwnProperty,o=Object.defineProperty||function(t,e,n){t[e]=n.value},i="function"==typeof Symbol?Symbol:{},a=i.iterator||"@@iterator",c=i.asyncIterator||"@@asyncIterator",s=i.toStringTag||"@@toStringTag";function u(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{u({},"")}catch(t){u=function(t,e,n){return t[e]=n}}function l(t,e,n,r){var i=e&&e.prototype instanceof y?e:y,a=Object.create(i.prototype),c=new _(r||[]);return o(a,"_invoke",{value:I(t,n,c)}),a}function d(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}t.wrap=l;var f="suspendedStart",h="suspendedYield",p="executing",v="completed",m={};function y(){}function g(){}function w(){}var b={};u(b,a,(function(){return this}));var L=Object.getPrototypeOf,x=L&&L(L(j([])));x&&x!==n&&r.call(x,a)&&(b=x);var E=w.prototype=y.prototype=Object.create(b);function O(t){["next","throw","return"].forEach((function(e){u(t,e,(function(t){return this._invoke(e,t)}))}))}function k(t,e){function n(o,i,a,c){var s=d(t[o],t,i);if("throw"!==s.type){var u=s.arg,l=u.value;return l&&"object"==typeof l&&r.call(l,"__await")?e.resolve(l.__await).then((function(t){n("next",t,a,c)}),(function(t){n("throw",t,a,c)})):e.resolve(l).then((function(t){u.value=t,a(u)}),(function(t){return n("throw",t,a,c)}))}c(s.arg)}var i;o(this,"_invoke",{value:function(t,r){function o(){return new e((function(e,o){n(t,r,e,o)}))}return i=i?i.then(o,o):o()}})}function I(t,e,n){var r=f;return function(o,i){if(r===p)throw new Error("Generator is already running");if(r===v){if("throw"===o)throw i;return T()}for(n.method=o,n.arg=i;;){var a=n.delegate;if(a){var c=P(a,n);if(c){if(c===m)continue;return c}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if(r===f)throw r=v,n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);r=p;var s=d(t,e,n);if("normal"===s.type){if(r=n.done?v:h,s.arg===m)continue;return{value:s.arg,done:n.done}}"throw"===s.type&&(r=v,n.method="throw",n.arg=s.arg)}}}function P(t,n){var r=n.method,o=t.iterator[r];if(o===e)return n.delegate=null,"throw"===r&&t.iterator.return&&(n.method="return",n.arg=e,P(t,n),"throw"===n.method)||"return"!==r&&(n.method="throw",n.arg=new TypeError("The iterator does not provide a '"+r+"' method")),m;var i=d(o,t.iterator,n.arg);if("throw"===i.type)return n.method="throw",n.arg=i.arg,n.delegate=null,m;var a=i.arg;return a?a.done?(n[t.resultName]=a.value,n.next=t.nextLoc,"return"!==n.method&&(n.method="next",n.arg=e),n.delegate=null,m):a:(n.method="throw",n.arg=new TypeError("iterator result is not an object"),n.delegate=null,m)}function A(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function B(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function _(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(A,this),this.reset(!0)}function j(t){if(t){var n=t[a];if(n)return n.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var o=-1,i=function n(){for(;++o<t.length;)if(r.call(t,o))return n.value=t[o],n.done=!1,n;return n.value=e,n.done=!0,n};return i.next=i}}return{next:T}}function T(){return{value:e,done:!0}}return g.prototype=w,o(E,"constructor",{value:w,configurable:!0}),o(w,"constructor",{value:g,configurable:!0}),g.displayName=u(w,s,"GeneratorFunction"),t.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===g||"GeneratorFunction"===(e.displayName||e.name))},t.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,w):(t.__proto__=w,u(t,s,"GeneratorFunction")),t.prototype=Object.create(E),t},t.awrap=function(t){return{__await:t}},O(k.prototype),u(k.prototype,c,(function(){return this})),t.AsyncIterator=k,t.async=function(e,n,r,o,i){void 0===i&&(i=Promise);var a=new k(l(e,n,r,o),i);return t.isGeneratorFunction(n)?a:a.next().then((function(t){return t.done?t.value:a.next()}))},O(E),u(E,s,"Generator"),u(E,a,(function(){return this})),u(E,"toString",(function(){return"[object Generator]"})),t.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},t.values=j,_.prototype={constructor:_,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=e,this.done=!1,this.delegate=null,this.method="next",this.arg=e,this.tryEntries.forEach(B),!t)for(var n in this)"t"===n.charAt(0)&&r.call(this,n)&&!isNaN(+n.slice(1))&&(this[n]=e)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var n=this;function o(r,o){return c.type="throw",c.arg=t,n.next=r,o&&(n.method="next",n.arg=e),!!o}for(var i=this.tryEntries.length-1;i>=0;--i){var a=this.tryEntries[i],c=a.completion;if("root"===a.tryLoc)return o("end");if(a.tryLoc<=this.prev){var s=r.call(a,"catchLoc"),u=r.call(a,"finallyLoc");if(s&&u){if(this.prev<a.catchLoc)return o(a.catchLoc,!0);if(this.prev<a.finallyLoc)return o(a.finallyLoc)}else if(s){if(this.prev<a.catchLoc)return o(a.catchLoc,!0)}else{if(!u)throw new Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return o(a.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var o=this.tryEntries[n];if(o.tryLoc<=this.prev&&r.call(o,"finallyLoc")&&this.prev<o.finallyLoc){var i=o;break}}i&&("break"===t||"continue"===t)&&i.tryLoc<=e&&e<=i.finallyLoc&&(i=null);var a=i?i.completion:{};return a.type=t,a.arg=e,i?(this.method="next",this.next=i.finallyLoc,m):this.complete(a)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),m},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),B(n),m}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;B(n)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,n,r){return this.delegate={iterator:j(t),resultName:n,nextLoc:r},"next"===this.method&&(this.arg=e),m}},t}(t.exports);try{regeneratorRuntime=e}catch(t){"object"==typeof globalThis?globalThis.regeneratorRuntime=e:Function("r","regeneratorRuntime = r")(e)}},44944:function(t,e,n){"use strict";t.exports=n.p+"assets/logo-filled.png"},60806:function(t,e,n){"use strict";t.exports=n.p+"ea00fd3244bdee7f9ff0.css"}},e={};function n(r){var o=e[r];if(void 0!==o)return o.exports;var i=e[r]={exports:{}};return t[r](i,i.exports,n),i.exports}n.m=t,n.n=function(t){var e=t&&t.__esModule?function(){return t.default}:function(){return t};return n.d(e,{a:e}),e},n.d=function(t,e){for(var r in e)n.o(e,r)&&!n.o(t,r)&&Object.defineProperty(t,r,{enumerable:!0,get:e[r]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(t){if("object"==typeof window)return window}}(),n.o=function(t,e){return Object.prototype.hasOwnProperty.call(t,e)},function(){var t;n.g.importScripts&&(t=n.g.location+"");var e=n.g.document;if(!t&&e&&(e.currentScript&&(t=e.currentScript.src),!t)){var r=e.getElementsByTagName("script");if(r.length)for(var o=r.length-1;o>-1&&!t;)t=r[o--].src}if(!t)throw new Error("Automatic publicPath is not supported in this browser");t=t.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=t}(),n.b=document.baseURI||self.location.href,function(){"use strict";function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(){e=function(){return r};var n,r={},o=Object.prototype,i=o.hasOwnProperty,a=Object.defineProperty||function(t,e,n){t[e]=n.value},c="function"==typeof Symbol?Symbol:{},s=c.iterator||"@@iterator",u=c.asyncIterator||"@@asyncIterator",l=c.toStringTag||"@@toStringTag";function d(t,e,n){return Object.defineProperty(t,e,{value:n,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{d({},"")}catch(n){d=function(t,e,n){return t[e]=n}}function f(t,e,n,r){var o=e&&e.prototype instanceof w?e:w,i=Object.create(o.prototype),c=new T(r||[]);return a(i,"_invoke",{value:A(t,n,c)}),i}function h(t,e,n){try{return{type:"normal",arg:t.call(e,n)}}catch(t){return{type:"throw",arg:t}}}r.wrap=f;var p="suspendedStart",v="suspendedYield",m="executing",y="completed",g={};function w(){}function b(){}function L(){}var x={};d(x,s,(function(){return this}));var E=Object.getPrototypeOf,O=E&&E(E(C([])));O&&O!==o&&i.call(O,s)&&(x=O);var k=L.prototype=w.prototype=Object.create(x);function I(t){["next","throw","return"].forEach((function(e){d(t,e,(function(t){return this._invoke(e,t)}))}))}function P(e,n){function r(o,a,c,s){var u=h(e[o],e,a);if("throw"!==u.type){var l=u.arg,d=l.value;return d&&"object"==t(d)&&i.call(d,"__await")?n.resolve(d.__await).then((function(t){r("next",t,c,s)}),(function(t){r("throw",t,c,s)})):n.resolve(d).then((function(t){l.value=t,c(l)}),(function(t){return r("throw",t,c,s)}))}s(u.arg)}var o;a(this,"_invoke",{value:function(t,e){function i(){return new n((function(n,o){r(t,e,n,o)}))}return o=o?o.then(i,i):i()}})}function A(t,e,r){var o=p;return function(i,a){if(o===m)throw new Error("Generator is already running");if(o===y){if("throw"===i)throw a;return{value:n,done:!0}}for(r.method=i,r.arg=a;;){var c=r.delegate;if(c){var s=B(c,r);if(s){if(s===g)continue;return s}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if(o===p)throw o=y,r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);o=m;var u=h(t,e,r);if("normal"===u.type){if(o=r.done?y:v,u.arg===g)continue;return{value:u.arg,done:r.done}}"throw"===u.type&&(o=y,r.method="throw",r.arg=u.arg)}}}function B(t,e){var r=e.method,o=t.iterator[r];if(o===n)return e.delegate=null,"throw"===r&&t.iterator.return&&(e.method="return",e.arg=n,B(t,e),"throw"===e.method)||"return"!==r&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+r+"' method")),g;var i=h(o,t.iterator,e.arg);if("throw"===i.type)return e.method="throw",e.arg=i.arg,e.delegate=null,g;var a=i.arg;return a?a.done?(e[t.resultName]=a.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=n),e.delegate=null,g):a:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,g)}function _(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function j(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function T(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(_,this),this.reset(!0)}function C(e){if(e||""===e){var r=e[s];if(r)return r.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var o=-1,a=function t(){for(;++o<e.length;)if(i.call(e,o))return t.value=e[o],t.done=!1,t;return t.value=n,t.done=!0,t};return a.next=a}}throw new TypeError(t(e)+" is not iterable")}return b.prototype=L,a(k,"constructor",{value:L,configurable:!0}),a(L,"constructor",{value:b,configurable:!0}),b.displayName=d(L,l,"GeneratorFunction"),r.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===b||"GeneratorFunction"===(e.displayName||e.name))},r.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,L):(t.__proto__=L,d(t,l,"GeneratorFunction")),t.prototype=Object.create(k),t},r.awrap=function(t){return{__await:t}},I(P.prototype),d(P.prototype,u,(function(){return this})),r.AsyncIterator=P,r.async=function(t,e,n,o,i){void 0===i&&(i=Promise);var a=new P(f(t,e,n,o),i);return r.isGeneratorFunction(e)?a:a.next().then((function(t){return t.done?t.value:a.next()}))},I(k),d(k,l,"Generator"),d(k,s,(function(){return this})),d(k,"toString",(function(){return"[object Generator]"})),r.keys=function(t){var e=Object(t),n=[];for(var r in e)n.push(r);return n.reverse(),function t(){for(;n.length;){var r=n.pop();if(r in e)return t.value=r,t.done=!1,t}return t.done=!0,t}},r.values=C,T.prototype={constructor:T,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=n,this.done=!1,this.delegate=null,this.method="next",this.arg=n,this.tryEntries.forEach(j),!t)for(var e in this)"t"===e.charAt(0)&&i.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=n)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function r(r,o){return c.type="throw",c.arg=t,e.next=r,o&&(e.method="next",e.arg=n),!!o}for(var o=this.tryEntries.length-1;o>=0;--o){var a=this.tryEntries[o],c=a.completion;if("root"===a.tryLoc)return r("end");if(a.tryLoc<=this.prev){var s=i.call(a,"catchLoc"),u=i.call(a,"finallyLoc");if(s&&u){if(this.prev<a.catchLoc)return r(a.catchLoc,!0);if(this.prev<a.finallyLoc)return r(a.finallyLoc)}else if(s){if(this.prev<a.catchLoc)return r(a.catchLoc,!0)}else{if(!u)throw new Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return r(a.finallyLoc)}}}},abrupt:function(t,e){for(var n=this.tryEntries.length-1;n>=0;--n){var r=this.tryEntries[n];if(r.tryLoc<=this.prev&&i.call(r,"finallyLoc")&&this.prev<r.finallyLoc){var o=r;break}}o&&("break"===t||"continue"===t)&&o.tryLoc<=e&&e<=o.finallyLoc&&(o=null);var a=o?o.completion:{};return a.type=t,a.arg=e,o?(this.method="next",this.next=o.finallyLoc,g):this.complete(a)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),g},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.finallyLoc===t)return this.complete(n.completion,n.afterLoc),j(n),g}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var n=this.tryEntries[e];if(n.tryLoc===t){var r=n.completion;if("throw"===r.type){var o=r.arg;j(n)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,e,r){return this.delegate={iterator:C(t),resultName:e,nextLoc:r},"next"===this.method&&(this.arg=n),g}},r}function r(t,e,n,r,o,i,a){try{var c=t[i](a),s=c.value}catch(t){return void n(t)}c.done?e(s):Promise.resolve(s).then(r,o)}function o(t){return function(){var e=this,n=arguments;return new Promise((function(o,i){var a=t.apply(e,n);function c(t){r(a,o,i,c,s,"next",t)}function s(t){r(a,o,i,c,s,"throw",t)}c(void 0)}))}}n(35666),Office.onReady((function(t){t.host===Office.HostType.Word&&(document.getElementById("app-body").style.display="flex");var n,r,i={preview:{label:"Preview",action:"preview"},publish:{label:"Publish",id:"publish"},viewProduction:{label:"View Production",id:"viewProduction",icon:"ms-Icon--OpenInNewWindow"},viewLibrary:{label:"View Library",id:"viewLibrary",icon:"ms-Icon--Library"},editConfig:{label:"Edit Config",id:"editConfig",icon:"ms-Icon--Settings"}},a=Office.context.document.settings.get("aemRepo"),c=a.replace("https://github.com/",""),s=document.getElementById("first-run"),u=Office.context.document.settings.get("contentUrl"),l=document.getElementById("config"),d=document.getElementById("aemPage"),f=document.getElementById("pageMetadata"),h=document.getElementById("pageOptions"),p=Office.context.document.settings.get("productionUrl"),v=document.getElementById("loader"),m=document.getElementById("lastPublished"),y=document.getElementById("lastPreviewed"),g=document.getElementById("lastModified"),w=document.createElement("button");w.classList.add("ms-Button"),w.setAttribute("id","expandPageMetadata"),w.setAttribute("type","button"),w.setAttribute("name","expandPageMetadata"),w.innerHTML='<i id="pageMetadata-close" class="ms-Icon ms-Icon--ChevronDown ms-font-xl"></i>\n  ',w.addEventListener("click",(function(t){t.stopPropagation(),f.classList.toggle("expanded"),w.classList.toggle("expanded")}));var b=document.getElementById("pageMetadata-controls");for(var L in b.appendChild(w),i){var x=document.createElement("button");x.classList.add("ms-Button"),x.setAttribute("id",L),x.setAttribute("type","button"),x.setAttribute("name",L),i[L].icon?(x.classList.add("ms-Button-withIcon"),x.innerHTML='<span class="ms-Button-icon"><i class="ms-Icon '.concat(i[L].icon,'"></i></span><span class="ms-Button-label">').concat(i[L].label,"</span>")):x.innerHTML='<span class="ms-Button-label">'.concat(i[L].label,"</span>"),x.addEventListener("click",(function(t){t.stopPropagation();var e=t.currentTarget.getAttribute("id");E["".concat(e)]()})),document.getElementById("pageOptions").appendChild(x);var E={publish:function(){var t=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(r){var o;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(o=document.getElementById("publish")).setAttribute("disabled","disabled"),o.textContent="Publishing...",o.classList.add("disabled"),v.classList.remove("d-none"),fetch("https://admin.hlx.page/live/"+c+n,{method:"POST",body:null}).then((function(t){return t.json()})).then((function(t){O(),d.src="".concat(t.live.url,"?date=").concat(Date.now()),d.addEventListener("load",(function(){v.classList.add("d-none"),o.textContent="Publish",o.removeAttribute("disabled"),o.classList.remove("disabled")}),!0)})),t.next=9,r.sync();case 9:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)})));return function(){return t.apply(this,arguments)}}(),preview:function(){var t=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(r){var o;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(o=document.getElementById("preview")).textContent="Previewing...",o.setAttribute("disabled","disabled"),o.classList.add("disabled"),v.classList.remove("d-none"),fetch("https://admin.hlx.page/preview/"+c+n,{method:"POST"}).then((function(t){return t.json()})).then((function(t){O(),d.src="".concat(t.preview.url,"?date=").concat(Date.now()),d.addEventListener("load",(function(){v.classList.add("d-none"),o.textContent="Preview",o.removeAttribute("disabled"),o.classList.remove("disabled")}),!0)})),t.next=9,r.sync();case 9:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)})));return function(){return t.apply(this,arguments)}}(),unpublish:function(){var t=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(r){var o,i,a,s,u;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(o=document.createElement("div")).classList.add("modal"),o.classList.add("fade"),o.classList.add("show"),o.setAttribute("id","unpublishModal"),o.setAttribute("tabindex","-1"),o.setAttribute("role","dialog"),o.setAttribute("aria-labelledby","unpublishModalLabel"),o.setAttribute("aria-hidden","true"),(i=document.createElement("div")).classList.add("modal-content"),(a=document.createElement("div")).classList.add("modal-actions"),i.innerHTML="<h2>Are you sure you want to unpublish this content?</h2>\n          <p>Unpublishing content will make the page not visible for users</p>",(s=document.createElement("button")).classList.add("ms-Button-label"),s.setAttribute("id","unpublishConfirm"),s.textContent="Unpublish",(u=document.createElement("button")).classList.add("ms-Button-label"),u.setAttribute("id","unpublishCancel"),u.setAttribute("data-dismiss","modal"),u.textContent="Cancel",u.addEventListener("click",(function(){o.classList.remove("show"),o.setAttribute("aria-hidden","true"),o.setAttribute("style","display: none"),o.setAttribute("aria-modal","false")})),s.addEventListener("click",(function(){o.classList.remove("show"),o.setAttribute("aria-hidden","true"),o.setAttribute("style","display: none"),o.setAttribute("aria-modal","false"),fetch("https://admin.hlx.page/live/"+c+n,{method:"DELETE"}).then((function(t){return t.json()})).then((function(t){t.live.lastModified?g.innerHTML="Last modified: ".concat(t.live.lastModified):g.innerHTML="No page published yet",d.src="".concat(t.live.url,"?date=").concat(Date.now())}))})),a.appendChild(s),a.appendChild(u),i.appendChild(a),o.appendChild(i),document.body.appendChild(o),t.next=32,r.sync();case 32:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)})));return function(){return t.apply(this,arguments)}}(),viewProduction:function(){var t=o(e().mark((function t(){var n;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:p?(n=new URL(r),window.open("https://".concat(p+n.pathname),"_blank")):window.open(r,"_blank");case 1:case"end":return t.stop()}}),t)})));return function(){return t.apply(this,arguments)}}(),editConfig:function(){var t=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(n){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return s.classList.add("d-none"),a=Office.context.document.settings.get("aemRepo"),p=Office.context.document.settings.get("productionUrl"),u=Office.context.document.settings.get("contentUrl"),p&&(document.getElementById("productionUrl").value=p),document.getElementById("contentUrl").value=u,document.getElementById("aemRepo").value=a,document.getElementById("saveConfig").addEventListener("click",(function(){E.saveConfig()})),f.classList.add("d-none"),d.classList.add("d-none"),h.classList.add("d-none"),l.classList.remove("d-none"),t.next=15,n.sync();case 15:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)})));return function(){return t.apply(this,arguments)}}(),saveConfig:function(){var t=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(n){var r;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return(r=document.getElementById("config-error")).classList.add("d-none"),""!=a&&""!=u||(r.innerHTML="Please enter both Github repo and Content URL fields",r.classList.remove("d-none")),a=document.getElementById("aemRepo").value,p=document.getElementById("productionUrl").value,u=document.getElementById("contentUrl").value,Office.context.document.settings.set("aemRepo",a),Office.context.document.settings.set("productionUrl",p),Office.context.document.settings.set("contentUrl",u),Office.context.document.settings.saveAsync((function(t){t.status==Office.AsyncResultStatus.Failed?console.log("Settings save failed. Error: "+t.error.message):console.log("Settings saved.")})),E.checkConfig(),t.next=13,n.sync();case 13:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)})));return function(){return t.apply(this,arguments)}}(),checkConfig:function(){var t=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(n){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:a&&u?(E.getInitialState(a),l.classList.add("d-none"),d.classList.remove("d-none")):(l.classList.remove("d-none"),d.classList.add("d-none"));case 1:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)})));return function(){return t.apply(this,arguments)}}(),getInitialState:function(){var t=o(e().mark((function t(){return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.abrupt("return",Word.run(function(){var t=o(e().mark((function t(o){var i,s,l,d,f,h,v,m;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:n=Office.context.document.url.replace(" ","%20").replace(u,"").replace(/ /g,"-").replace(".docx","").replace(/’/g,"-").toLowerCase(),c=a.replace("https://github.com/",""),r="https://admin.hlx.page/live/"+c+n,p=Office.context.document.settings.get("productionUrl"),i=document.getElementById("preview"),s=document.getElementById("publish"),l=document.getElementById("pageMetadata"),d=document.getElementById("viewProduction"),f=document.getElementById("pageOptions"),h=document.getElementById("viewLibrary"),(v=document.getElementById("aemPage")).classList.add("d-none"),(m=document.createElement("div")).classList.add("small-loader"),m.setAttribute("id","loader"),m.innerHTML='<div id="loader" class="loader transparent">\n        <img width="50" height="50" style="margin-bottom:50px; margin-top: -150px;" src="../../assets/logo-filled.png" alt="AEM" title="AEM" />\n\n        <div class="lds-default"><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div></div>\n        <span>Please wait...</span>\n    </div>',document.body.appendChild(m),fetch("https://admin.hlx.page/status/"+c+n,{method:"GET"}).then((function(t){return t.json()})).then((function(t){if(v.src="".concat(t.preview.url,"?date=").concat(Date.now()),v.addEventListener("load",(function(){v.classList.remove("d-none"),m.classList.add("d-none"),i.textContent="Preview",s.textContent="Publish",l.classList.remove("d-none"),f.classList.remove("d-none")}),!0),O(),t.live.url){d.classList.remove("d-none"),r=t.live.url;var e=new URL(t.preview.url),n="https://".concat(e.hostname,"/tools/sidekick/library.html");fetch(n,{method:"GET"}).then((function(t){200==t.status&&(h.addEventListener("click",(function(t){t.stopPropagation(),window.open("https://".concat(e.hostname,"/tools/sidekick/library.html"),"_blank")})),h.classList.remove("d-none"))}))}}));case 19:case"end":return t.stop()}}),t)})));return function(e){return t.apply(this,arguments)}}()));case 1:case"end":return t.stop()}}),t)})));return function(){return t.apply(this,arguments)}}()}}function O(){var t="https://admin.hlx.page/status/"+c+n;console.log(t),fetch(t,{method:"GET"}).then((function(t){return t.json()})).then((function(t){console.log(Intl.DateTimeFormat().resolvedOptions().timeZone),console.log(t);var e=new Date(t.live.sourceLastModified).toLocaleString("en-AU",{day:"numeric",month:"long",year:"numeric",hour:"2-digit",minute:"2-digit",second:"2-digit"}),n=new Date(t.live.lastModified).toLocaleString("en-AU",{day:"numeric",month:"long",year:"numeric",hour:"2-digit",minute:"2-digit",second:"2-digit"}),r=new Date(t.preview.lastModified).toLocaleString("en-AU",{day:"numeric",month:"long",year:"numeric",hour:"2-digit",minute:"2-digit",second:"2-digit"});g.innerHTML="Last modified: ".concat(e),m.innerHTML="Last published: ".concat(n),y.innerHTML="Last previewed: ".concat(r)}))}E.checkConfig()}))}(),function(){"use strict";var t=n(27091),e=n.n(t),r=new URL(n(60806),n.b),o=new URL(n(44944),n.b);e()(r),e()(o)}()}();
//# sourceMappingURL=taskpane.js.map