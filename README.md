# XLSX Accessibility Checker

This is version of the accessibility checker, exported from Observable (https://observablehq.com/@jtrim-ons/xlsx-accessibility-checker).

It can be run on a local web server. (A few JavaScript libraries are fetched from cdn.jsdelivr.net).

## Notes on how I've modified this after exporting from Observable

I've tweaked index.html.  I also re-ordered some of the cells to put the styles early and named a cell "start_of_appendix".

## XLSX Accessibility Checker

https://observablehq.com/@jtrim-ons/xlsx-accessibility-checker@822

View this notebook in your browser by running a web server in this folder. For
example:

~~~sh
npx http-server
~~~

Or, use the [Observable Runtime](https://github.com/observablehq/runtime) to
import this module directly into your application. To npm install:

~~~sh
npm install @observablehq/runtime@5
npm install https://api.observablehq.com/@jtrim-ons/xlsx-accessibility-checker@822.tgz?v=3
~~~

Then, import your notebook and the runtime as:

~~~js
import {Runtime, Inspector} from "@observablehq/runtime";
import define from "@jtrim-ons/xlsx-accessibility-checker";
~~~

To log the value of the cell named “foo”:

~~~js
const runtime = new Runtime();
const main = runtime.module(define);
main.value("foo").then(value => console.log(value));
~~~
