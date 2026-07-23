# Legacy DOC reader notice

`word-extractor-browser.js` is a browser-targeted, OLE-only bundle derived from
[`word-extractor` 1.0.4](https://github.com/morungos/node-word-extractor) under
the MIT license in `word-extractor.LICENSE.txt`.

The bundle contains two compatibility adjustments required by the browser
polyfills:

- OLE directory names exclude the complete two-byte UTF-16 null terminator.
- OLE storage sectors are emitted by a small asynchronous event pump instead
  of Node's native `Readable` stream.

It only extracts text from classic Word 97–2003 OLE documents. It does not
execute macros, scripts, embedded objects, or links. Import remains best-effort
and every extracted field is shown for user review before it can change a
rider.
