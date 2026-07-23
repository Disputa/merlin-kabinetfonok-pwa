(function(f){if(typeof exports==="object"&&typeof module!=="undefined"){module.exports=f()}else if(typeof define==="function"&&define.amd){define([],f)}else{var g;if(typeof window!=="undefined"){g=window}else if(typeof global!=="undefined"){g=global}else if(typeof self!=="undefined"){g=self}else{g=this}g.LegacyDocReader = f()}})(function(){var define,module,exports;return (function(){function r(e,n,t){function o(i,f){if(!n[i]){if(!e[i]){var c="function"==typeof require&&require;if(!f&&c)return c(i,!0);if(u)return u(i,!0);var a=new Error("Cannot find module '"+i+"'");throw a.code="MODULE_NOT_FOUND",a}var p=n[i]={exports:{}};e[i][0].call(p.exports,function(r){var n=e[i][1][r];return o(n||r)},p,p.exports,r,e,n,t)}return n[i].exports}for(var u="function"==typeof require&&require,i=0;i<t.length;i++)o(t[i]);return o}return r})()({1:[function(require,module,exports){
const { Buffer } = require("buffer");
const BufferReader = require("word-extractor/lib/buffer-reader");
const WordOleExtractor = require("word-extractor/lib/word-ole-extractor");

module.exports = {
  async extract(arrayBuffer) {
    const reader = new BufferReader(Buffer.from(arrayBuffer));
    await reader.open();
    const document = await new WordOleExtractor().extract(reader);
    return [
      document.getHeaders(),
      document.getBody(),
      document.getFootnotes(),
      document.getEndnotes(),
      document.getTextboxes(),
    ]
      .filter(Boolean)
      .join("\n");
  },
};

},{"buffer":3,"word-extractor/lib/buffer-reader":6,"word-extractor/lib/word-ole-extractor":15}],2:[function(require,module,exports){
'use strict'

exports.byteLength = byteLength
exports.toByteArray = toByteArray
exports.fromByteArray = fromByteArray

var lookup = []
var revLookup = []
var Arr = typeof Uint8Array !== 'undefined' ? Uint8Array : Array

var code = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'
for (var i = 0, len = code.length; i < len; ++i) {
  lookup[i] = code[i]
  revLookup[code.charCodeAt(i)] = i
}

// Support decoding URL-safe base64 strings, as Node.js does.
// See: https://en.wikipedia.org/wiki/Base64#URL_applications
revLookup['-'.charCodeAt(0)] = 62
revLookup['_'.charCodeAt(0)] = 63

function getLens (b64) {
  var len = b64.length

  if (len % 4 > 0) {
    throw new Error('Invalid string. Length must be a multiple of 4')
  }

  // Trim off extra bytes after placeholder bytes are found
  // See: https://github.com/beatgammit/base64-js/issues/42
  var validLen = b64.indexOf('=')
  if (validLen === -1) validLen = len

  var placeHoldersLen = validLen === len
    ? 0
    : 4 - (validLen % 4)

  return [validLen, placeHoldersLen]
}

// base64 is 4/3 + up to two characters of the original data
function byteLength (b64) {
  var lens = getLens(b64)
  var validLen = lens[0]
  var placeHoldersLen = lens[1]
  return ((validLen + placeHoldersLen) * 3 / 4) - placeHoldersLen
}

function _byteLength (b64, validLen, placeHoldersLen) {
  return ((validLen + placeHoldersLen) * 3 / 4) - placeHoldersLen
}

function toByteArray (b64) {
  var tmp
  var lens = getLens(b64)
  var validLen = lens[0]
  var placeHoldersLen = lens[1]

  var arr = new Arr(_byteLength(b64, validLen, placeHoldersLen))

  var curByte = 0

  // if there are placeholders, only get up to the last complete 4 chars
  var len = placeHoldersLen > 0
    ? validLen - 4
    : validLen

  var i
  for (i = 0; i < len; i += 4) {
    tmp =
      (revLookup[b64.charCodeAt(i)] << 18) |
      (revLookup[b64.charCodeAt(i + 1)] << 12) |
      (revLookup[b64.charCodeAt(i + 2)] << 6) |
      revLookup[b64.charCodeAt(i + 3)]
    arr[curByte++] = (tmp >> 16) & 0xFF
    arr[curByte++] = (tmp >> 8) & 0xFF
    arr[curByte++] = tmp & 0xFF
  }

  if (placeHoldersLen === 2) {
    tmp =
      (revLookup[b64.charCodeAt(i)] << 2) |
      (revLookup[b64.charCodeAt(i + 1)] >> 4)
    arr[curByte++] = tmp & 0xFF
  }

  if (placeHoldersLen === 1) {
    tmp =
      (revLookup[b64.charCodeAt(i)] << 10) |
      (revLookup[b64.charCodeAt(i + 1)] << 4) |
      (revLookup[b64.charCodeAt(i + 2)] >> 2)
    arr[curByte++] = (tmp >> 8) & 0xFF
    arr[curByte++] = tmp & 0xFF
  }

  return arr
}

function tripletToBase64 (num) {
  return lookup[num >> 18 & 0x3F] +
    lookup[num >> 12 & 0x3F] +
    lookup[num >> 6 & 0x3F] +
    lookup[num & 0x3F]
}

function encodeChunk (uint8, start, end) {
  var tmp
  var output = []
  for (var i = start; i < end; i += 3) {
    tmp =
      ((uint8[i] << 16) & 0xFF0000) +
      ((uint8[i + 1] << 8) & 0xFF00) +
      (uint8[i + 2] & 0xFF)
    output.push(tripletToBase64(tmp))
  }
  return output.join('')
}

function fromByteArray (uint8) {
  var tmp
  var len = uint8.length
  var extraBytes = len % 3 // if we have 1 byte left, pad 2 bytes
  var parts = []
  var maxChunkLength = 16383 // must be multiple of 3

  // go through the array every three bytes, we'll deal with trailing stuff later
  for (var i = 0, len2 = len - extraBytes; i < len2; i += maxChunkLength) {
    parts.push(encodeChunk(uint8, i, (i + maxChunkLength) > len2 ? len2 : (i + maxChunkLength)))
  }

  // pad the end with zeros, but make sure to not forget the extra bytes
  if (extraBytes === 1) {
    tmp = uint8[len - 1]
    parts.push(
      lookup[tmp >> 2] +
      lookup[(tmp << 4) & 0x3F] +
      '=='
    )
  } else if (extraBytes === 2) {
    tmp = (uint8[len - 2] << 8) + uint8[len - 1]
    parts.push(
      lookup[tmp >> 10] +
      lookup[(tmp >> 4) & 0x3F] +
      lookup[(tmp << 2) & 0x3F] +
      '='
    )
  }

  return parts.join('')
}

},{}],3:[function(require,module,exports){
(function (Buffer){(function (){
/*!
 * The buffer module from node.js, for the browser.
 *
 * @author   Feross Aboukhadijeh <https://feross.org>
 * @license  MIT
 */
/* eslint-disable no-proto */

'use strict'

var base64 = require('base64-js')
var ieee754 = require('ieee754')

exports.Buffer = Buffer
exports.SlowBuffer = SlowBuffer
exports.INSPECT_MAX_BYTES = 50

var K_MAX_LENGTH = 0x7fffffff
exports.kMaxLength = K_MAX_LENGTH

/**
 * If `Buffer.TYPED_ARRAY_SUPPORT`:
 *   === true    Use Uint8Array implementation (fastest)
 *   === false   Print warning and recommend using `buffer` v4.x which has an Object
 *               implementation (most compatible, even IE6)
 *
 * Browsers that support typed arrays are IE 10+, Firefox 4+, Chrome 7+, Safari 5.1+,
 * Opera 11.6+, iOS 4.2+.
 *
 * We report that the browser does not support typed arrays if the are not subclassable
 * using __proto__. Firefox 4-29 lacks support for adding new properties to `Uint8Array`
 * (See: https://bugzilla.mozilla.org/show_bug.cgi?id=695438). IE 10 lacks support
 * for __proto__ and has a buggy typed array implementation.
 */
Buffer.TYPED_ARRAY_SUPPORT = typedArraySupport()

if (!Buffer.TYPED_ARRAY_SUPPORT && typeof console !== 'undefined' &&
    typeof console.error === 'function') {
  console.error(
    'This browser lacks typed array (Uint8Array) support which is required by ' +
    '`buffer` v5.x. Use `buffer` v4.x if you require old browser support.'
  )
}

function typedArraySupport () {
  // Can typed array instances can be augmented?
  try {
    var arr = new Uint8Array(1)
    arr.__proto__ = { __proto__: Uint8Array.prototype, foo: function () { return 42 } }
    return arr.foo() === 42
  } catch (e) {
    return false
  }
}

Object.defineProperty(Buffer.prototype, 'parent', {
  enumerable: true,
  get: function () {
    if (!Buffer.isBuffer(this)) return undefined
    return this.buffer
  }
})

Object.defineProperty(Buffer.prototype, 'offset', {
  enumerable: true,
  get: function () {
    if (!Buffer.isBuffer(this)) return undefined
    return this.byteOffset
  }
})

function createBuffer (length) {
  if (length > K_MAX_LENGTH) {
    throw new RangeError('The value "' + length + '" is invalid for option "size"')
  }
  // Return an augmented `Uint8Array` instance
  var buf = new Uint8Array(length)
  buf.__proto__ = Buffer.prototype
  return buf
}

/**
 * The Buffer constructor returns instances of `Uint8Array` that have their
 * prototype changed to `Buffer.prototype`. Furthermore, `Buffer` is a subclass of
 * `Uint8Array`, so the returned instances will have all the node `Buffer` methods
 * and the `Uint8Array` methods. Square bracket notation works as expected -- it
 * returns a single octet.
 *
 * The `Uint8Array` prototype remains unmodified.
 */

function Buffer (arg, encodingOrOffset, length) {
  // Common case.
  if (typeof arg === 'number') {
    if (typeof encodingOrOffset === 'string') {
      throw new TypeError(
        'The "string" argument must be of type string. Received type number'
      )
    }
    return allocUnsafe(arg)
  }
  return from(arg, encodingOrOffset, length)
}

// Fix subarray() in ES2016. See: https://github.com/feross/buffer/pull/97
if (typeof Symbol !== 'undefined' && Symbol.species != null &&
    Buffer[Symbol.species] === Buffer) {
  Object.defineProperty(Buffer, Symbol.species, {
    value: null,
    configurable: true,
    enumerable: false,
    writable: false
  })
}

Buffer.poolSize = 8192 // not used by this implementation

function from (value, encodingOrOffset, length) {
  if (typeof value === 'string') {
    return fromString(value, encodingOrOffset)
  }

  if (ArrayBuffer.isView(value)) {
    return fromArrayLike(value)
  }

  if (value == null) {
    throw TypeError(
      'The first argument must be one of type string, Buffer, ArrayBuffer, Array, ' +
      'or Array-like Object. Received type ' + (typeof value)
    )
  }

  if (isInstance(value, ArrayBuffer) ||
      (value && isInstance(value.buffer, ArrayBuffer))) {
    return fromArrayBuffer(value, encodingOrOffset, length)
  }

  if (typeof value === 'number') {
    throw new TypeError(
      'The "value" argument must not be of type number. Received type number'
    )
  }

  var valueOf = value.valueOf && value.valueOf()
  if (valueOf != null && valueOf !== value) {
    return Buffer.from(valueOf, encodingOrOffset, length)
  }

  var b = fromObject(value)
  if (b) return b

  if (typeof Symbol !== 'undefined' && Symbol.toPrimitive != null &&
      typeof value[Symbol.toPrimitive] === 'function') {
    return Buffer.from(
      value[Symbol.toPrimitive]('string'), encodingOrOffset, length
    )
  }

  throw new TypeError(
    'The first argument must be one of type string, Buffer, ArrayBuffer, Array, ' +
    'or Array-like Object. Received type ' + (typeof value)
  )
}

/**
 * Functionally equivalent to Buffer(arg, encoding) but throws a TypeError
 * if value is a number.
 * Buffer.from(str[, encoding])
 * Buffer.from(array)
 * Buffer.from(buffer)
 * Buffer.from(arrayBuffer[, byteOffset[, length]])
 **/
Buffer.from = function (value, encodingOrOffset, length) {
  return from(value, encodingOrOffset, length)
}

// Note: Change prototype *after* Buffer.from is defined to workaround Chrome bug:
// https://github.com/feross/buffer/pull/148
Buffer.prototype.__proto__ = Uint8Array.prototype
Buffer.__proto__ = Uint8Array

function assertSize (size) {
  if (typeof size !== 'number') {
    throw new TypeError('"size" argument must be of type number')
  } else if (size < 0) {
    throw new RangeError('The value "' + size + '" is invalid for option "size"')
  }
}

function alloc (size, fill, encoding) {
  assertSize(size)
  if (size <= 0) {
    return createBuffer(size)
  }
  if (fill !== undefined) {
    // Only pay attention to encoding if it's a string. This
    // prevents accidentally sending in a number that would
    // be interpretted as a start offset.
    return typeof encoding === 'string'
      ? createBuffer(size).fill(fill, encoding)
      : createBuffer(size).fill(fill)
  }
  return createBuffer(size)
}

/**
 * Creates a new filled Buffer instance.
 * alloc(size[, fill[, encoding]])
 **/
Buffer.alloc = function (size, fill, encoding) {
  return alloc(size, fill, encoding)
}

function allocUnsafe (size) {
  assertSize(size)
  return createBuffer(size < 0 ? 0 : checked(size) | 0)
}

/**
 * Equivalent to Buffer(num), by default creates a non-zero-filled Buffer instance.
 * */
Buffer.allocUnsafe = function (size) {
  return allocUnsafe(size)
}
/**
 * Equivalent to SlowBuffer(num), by default creates a non-zero-filled Buffer instance.
 */
Buffer.allocUnsafeSlow = function (size) {
  return allocUnsafe(size)
}

function fromString (string, encoding) {
  if (typeof encoding !== 'string' || encoding === '') {
    encoding = 'utf8'
  }

  if (!Buffer.isEncoding(encoding)) {
    throw new TypeError('Unknown encoding: ' + encoding)
  }

  var length = byteLength(string, encoding) | 0
  var buf = createBuffer(length)

  var actual = buf.write(string, encoding)

  if (actual !== length) {
    // Writing a hex string, for example, that contains invalid characters will
    // cause everything after the first invalid character to be ignored. (e.g.
    // 'abxxcd' will be treated as 'ab')
    buf = buf.slice(0, actual)
  }

  return buf
}

function fromArrayLike (array) {
  var length = array.length < 0 ? 0 : checked(array.length) | 0
  var buf = createBuffer(length)
  for (var i = 0; i < length; i += 1) {
    buf[i] = array[i] & 255
  }
  return buf
}

function fromArrayBuffer (array, byteOffset, length) {
  if (byteOffset < 0 || array.byteLength < byteOffset) {
    throw new RangeError('"offset" is outside of buffer bounds')
  }

  if (array.byteLength < byteOffset + (length || 0)) {
    throw new RangeError('"length" is outside of buffer bounds')
  }

  var buf
  if (byteOffset === undefined && length === undefined) {
    buf = new Uint8Array(array)
  } else if (length === undefined) {
    buf = new Uint8Array(array, byteOffset)
  } else {
    buf = new Uint8Array(array, byteOffset, length)
  }

  // Return an augmented `Uint8Array` instance
  buf.__proto__ = Buffer.prototype
  return buf
}

function fromObject (obj) {
  if (Buffer.isBuffer(obj)) {
    var len = checked(obj.length) | 0
    var buf = createBuffer(len)

    if (buf.length === 0) {
      return buf
    }

    obj.copy(buf, 0, 0, len)
    return buf
  }

  if (obj.length !== undefined) {
    if (typeof obj.length !== 'number' || numberIsNaN(obj.length)) {
      return createBuffer(0)
    }
    return fromArrayLike(obj)
  }

  if (obj.type === 'Buffer' && Array.isArray(obj.data)) {
    return fromArrayLike(obj.data)
  }
}

function checked (length) {
  // Note: cannot use `length < K_MAX_LENGTH` here because that fails when
  // length is NaN (which is otherwise coerced to zero.)
  if (length >= K_MAX_LENGTH) {
    throw new RangeError('Attempt to allocate Buffer larger than maximum ' +
                         'size: 0x' + K_MAX_LENGTH.toString(16) + ' bytes')
  }
  return length | 0
}

function SlowBuffer (length) {
  if (+length != length) { // eslint-disable-line eqeqeq
    length = 0
  }
  return Buffer.alloc(+length)
}

Buffer.isBuffer = function isBuffer (b) {
  return b != null && b._isBuffer === true &&
    b !== Buffer.prototype // so Buffer.isBuffer(Buffer.prototype) will be false
}

Buffer.compare = function compare (a, b) {
  if (isInstance(a, Uint8Array)) a = Buffer.from(a, a.offset, a.byteLength)
  if (isInstance(b, Uint8Array)) b = Buffer.from(b, b.offset, b.byteLength)
  if (!Buffer.isBuffer(a) || !Buffer.isBuffer(b)) {
    throw new TypeError(
      'The "buf1", "buf2" arguments must be one of type Buffer or Uint8Array'
    )
  }

  if (a === b) return 0

  var x = a.length
  var y = b.length

  for (var i = 0, len = Math.min(x, y); i < len; ++i) {
    if (a[i] !== b[i]) {
      x = a[i]
      y = b[i]
      break
    }
  }

  if (x < y) return -1
  if (y < x) return 1
  return 0
}

Buffer.isEncoding = function isEncoding (encoding) {
  switch (String(encoding).toLowerCase()) {
    case 'hex':
    case 'utf8':
    case 'utf-8':
    case 'ascii':
    case 'latin1':
    case 'binary':
    case 'base64':
    case 'ucs2':
    case 'ucs-2':
    case 'utf16le':
    case 'utf-16le':
      return true
    default:
      return false
  }
}

Buffer.concat = function concat (list, length) {
  if (!Array.isArray(list)) {
    throw new TypeError('"list" argument must be an Array of Buffers')
  }

  if (list.length === 0) {
    return Buffer.alloc(0)
  }

  var i
  if (length === undefined) {
    length = 0
    for (i = 0; i < list.length; ++i) {
      length += list[i].length
    }
  }

  var buffer = Buffer.allocUnsafe(length)
  var pos = 0
  for (i = 0; i < list.length; ++i) {
    var buf = list[i]
    if (isInstance(buf, Uint8Array)) {
      buf = Buffer.from(buf)
    }
    if (!Buffer.isBuffer(buf)) {
      throw new TypeError('"list" argument must be an Array of Buffers')
    }
    buf.copy(buffer, pos)
    pos += buf.length
  }
  return buffer
}

function byteLength (string, encoding) {
  if (Buffer.isBuffer(string)) {
    return string.length
  }
  if (ArrayBuffer.isView(string) || isInstance(string, ArrayBuffer)) {
    return string.byteLength
  }
  if (typeof string !== 'string') {
    throw new TypeError(
      'The "string" argument must be one of type string, Buffer, or ArrayBuffer. ' +
      'Received type ' + typeof string
    )
  }

  var len = string.length
  var mustMatch = (arguments.length > 2 && arguments[2] === true)
  if (!mustMatch && len === 0) return 0

  // Use a for loop to avoid recursion
  var loweredCase = false
  for (;;) {
    switch (encoding) {
      case 'ascii':
      case 'latin1':
      case 'binary':
        return len
      case 'utf8':
      case 'utf-8':
        return utf8ToBytes(string).length
      case 'ucs2':
      case 'ucs-2':
      case 'utf16le':
      case 'utf-16le':
        return len * 2
      case 'hex':
        return len >>> 1
      case 'base64':
        return base64ToBytes(string).length
      default:
        if (loweredCase) {
          return mustMatch ? -1 : utf8ToBytes(string).length // assume utf8
        }
        encoding = ('' + encoding).toLowerCase()
        loweredCase = true
    }
  }
}
Buffer.byteLength = byteLength

function slowToString (encoding, start, end) {
  var loweredCase = false

  // No need to verify that "this.length <= MAX_UINT32" since it's a read-only
  // property of a typed array.

  // This behaves neither like String nor Uint8Array in that we set start/end
  // to their upper/lower bounds if the value passed is out of range.
  // undefined is handled specially as per ECMA-262 6th Edition,
  // Section 13.3.3.7 Runtime Semantics: KeyedBindingInitialization.
  if (start === undefined || start < 0) {
    start = 0
  }
  // Return early if start > this.length. Done here to prevent potential uint32
  // coercion fail below.
  if (start > this.length) {
    return ''
  }

  if (end === undefined || end > this.length) {
    end = this.length
  }

  if (end <= 0) {
    return ''
  }

  // Force coersion to uint32. This will also coerce falsey/NaN values to 0.
  end >>>= 0
  start >>>= 0

  if (end <= start) {
    return ''
  }

  if (!encoding) encoding = 'utf8'

  while (true) {
    switch (encoding) {
      case 'hex':
        return hexSlice(this, start, end)

      case 'utf8':
      case 'utf-8':
        return utf8Slice(this, start, end)

      case 'ascii':
        return asciiSlice(this, start, end)

      case 'latin1':
      case 'binary':
        return latin1Slice(this, start, end)

      case 'base64':
        return base64Slice(this, start, end)

      case 'ucs2':
      case 'ucs-2':
      case 'utf16le':
      case 'utf-16le':
        return utf16leSlice(this, start, end)

      default:
        if (loweredCase) throw new TypeError('Unknown encoding: ' + encoding)
        encoding = (encoding + '').toLowerCase()
        loweredCase = true
    }
  }
}

// This property is used by `Buffer.isBuffer` (and the `is-buffer` npm package)
// to detect a Buffer instance. It's not possible to use `instanceof Buffer`
// reliably in a browserify context because there could be multiple different
// copies of the 'buffer' package in use. This method works even for Buffer
// instances that were created from another copy of the `buffer` package.
// See: https://github.com/feross/buffer/issues/154
Buffer.prototype._isBuffer = true

function swap (b, n, m) {
  var i = b[n]
  b[n] = b[m]
  b[m] = i
}

Buffer.prototype.swap16 = function swap16 () {
  var len = this.length
  if (len % 2 !== 0) {
    throw new RangeError('Buffer size must be a multiple of 16-bits')
  }
  for (var i = 0; i < len; i += 2) {
    swap(this, i, i + 1)
  }
  return this
}

Buffer.prototype.swap32 = function swap32 () {
  var len = this.length
  if (len % 4 !== 0) {
    throw new RangeError('Buffer size must be a multiple of 32-bits')
  }
  for (var i = 0; i < len; i += 4) {
    swap(this, i, i + 3)
    swap(this, i + 1, i + 2)
  }
  return this
}

Buffer.prototype.swap64 = function swap64 () {
  var len = this.length
  if (len % 8 !== 0) {
    throw new RangeError('Buffer size must be a multiple of 64-bits')
  }
  for (var i = 0; i < len; i += 8) {
    swap(this, i, i + 7)
    swap(this, i + 1, i + 6)
    swap(this, i + 2, i + 5)
    swap(this, i + 3, i + 4)
  }
  return this
}

Buffer.prototype.toString = function toString () {
  var length = this.length
  if (length === 0) return ''
  if (arguments.length === 0) return utf8Slice(this, 0, length)
  return slowToString.apply(this, arguments)
}

Buffer.prototype.toLocaleString = Buffer.prototype.toString

Buffer.prototype.equals = function equals (b) {
  if (!Buffer.isBuffer(b)) throw new TypeError('Argument must be a Buffer')
  if (this === b) return true
  return Buffer.compare(this, b) === 0
}

Buffer.prototype.inspect = function inspect () {
  var str = ''
  var max = exports.INSPECT_MAX_BYTES
  str = this.toString('hex', 0, max).replace(/(.{2})/g, '$1 ').trim()
  if (this.length > max) str += ' ... '
  return '<Buffer ' + str + '>'
}

Buffer.prototype.compare = function compare (target, start, end, thisStart, thisEnd) {
  if (isInstance(target, Uint8Array)) {
    target = Buffer.from(target, target.offset, target.byteLength)
  }
  if (!Buffer.isBuffer(target)) {
    throw new TypeError(
      'The "target" argument must be one of type Buffer or Uint8Array. ' +
      'Received type ' + (typeof target)
    )
  }

  if (start === undefined) {
    start = 0
  }
  if (end === undefined) {
    end = target ? target.length : 0
  }
  if (thisStart === undefined) {
    thisStart = 0
  }
  if (thisEnd === undefined) {
    thisEnd = this.length
  }

  if (start < 0 || end > target.length || thisStart < 0 || thisEnd > this.length) {
    throw new RangeError('out of range index')
  }

  if (thisStart >= thisEnd && start >= end) {
    return 0
  }
  if (thisStart >= thisEnd) {
    return -1
  }
  if (start >= end) {
    return 1
  }

  start >>>= 0
  end >>>= 0
  thisStart >>>= 0
  thisEnd >>>= 0

  if (this === target) return 0

  var x = thisEnd - thisStart
  var y = end - start
  var len = Math.min(x, y)

  var thisCopy = this.slice(thisStart, thisEnd)
  var targetCopy = target.slice(start, end)

  for (var i = 0; i < len; ++i) {
    if (thisCopy[i] !== targetCopy[i]) {
      x = thisCopy[i]
      y = targetCopy[i]
      break
    }
  }

  if (x < y) return -1
  if (y < x) return 1
  return 0
}

// Finds either the first index of `val` in `buffer` at offset >= `byteOffset`,
// OR the last index of `val` in `buffer` at offset <= `byteOffset`.
//
// Arguments:
// - buffer - a Buffer to search
// - val - a string, Buffer, or number
// - byteOffset - an index into `buffer`; will be clamped to an int32
// - encoding - an optional encoding, relevant is val is a string
// - dir - true for indexOf, false for lastIndexOf
function bidirectionalIndexOf (buffer, val, byteOffset, encoding, dir) {
  // Empty buffer means no match
  if (buffer.length === 0) return -1

  // Normalize byteOffset
  if (typeof byteOffset === 'string') {
    encoding = byteOffset
    byteOffset = 0
  } else if (byteOffset > 0x7fffffff) {
    byteOffset = 0x7fffffff
  } else if (byteOffset < -0x80000000) {
    byteOffset = -0x80000000
  }
  byteOffset = +byteOffset // Coerce to Number.
  if (numberIsNaN(byteOffset)) {
    // byteOffset: it it's undefined, null, NaN, "foo", etc, search whole buffer
    byteOffset = dir ? 0 : (buffer.length - 1)
  }

  // Normalize byteOffset: negative offsets start from the end of the buffer
  if (byteOffset < 0) byteOffset = buffer.length + byteOffset
  if (byteOffset >= buffer.length) {
    if (dir) return -1
    else byteOffset = buffer.length - 1
  } else if (byteOffset < 0) {
    if (dir) byteOffset = 0
    else return -1
  }

  // Normalize val
  if (typeof val === 'string') {
    val = Buffer.from(val, encoding)
  }

  // Finally, search either indexOf (if dir is true) or lastIndexOf
  if (Buffer.isBuffer(val)) {
    // Special case: looking for empty string/buffer always fails
    if (val.length === 0) {
      return -1
    }
    return arrayIndexOf(buffer, val, byteOffset, encoding, dir)
  } else if (typeof val === 'number') {
    val = val & 0xFF // Search for a byte value [0-255]
    if (typeof Uint8Array.prototype.indexOf === 'function') {
      if (dir) {
        return Uint8Array.prototype.indexOf.call(buffer, val, byteOffset)
      } else {
        return Uint8Array.prototype.lastIndexOf.call(buffer, val, byteOffset)
      }
    }
    return arrayIndexOf(buffer, [ val ], byteOffset, encoding, dir)
  }

  throw new TypeError('val must be string, number or Buffer')
}

function arrayIndexOf (arr, val, byteOffset, encoding, dir) {
  var indexSize = 1
  var arrLength = arr.length
  var valLength = val.length

  if (encoding !== undefined) {
    encoding = String(encoding).toLowerCase()
    if (encoding === 'ucs2' || encoding === 'ucs-2' ||
        encoding === 'utf16le' || encoding === 'utf-16le') {
      if (arr.length < 2 || val.length < 2) {
        return -1
      }
      indexSize = 2
      arrLength /= 2
      valLength /= 2
      byteOffset /= 2
    }
  }

  function read (buf, i) {
    if (indexSize === 1) {
      return buf[i]
    } else {
      return buf.readUInt16BE(i * indexSize)
    }
  }

  var i
  if (dir) {
    var foundIndex = -1
    for (i = byteOffset; i < arrLength; i++) {
      if (read(arr, i) === read(val, foundIndex === -1 ? 0 : i - foundIndex)) {
        if (foundIndex === -1) foundIndex = i
        if (i - foundIndex + 1 === valLength) return foundIndex * indexSize
      } else {
        if (foundIndex !== -1) i -= i - foundIndex
        foundIndex = -1
      }
    }
  } else {
    if (byteOffset + valLength > arrLength) byteOffset = arrLength - valLength
    for (i = byteOffset; i >= 0; i--) {
      var found = true
      for (var j = 0; j < valLength; j++) {
        if (read(arr, i + j) !== read(val, j)) {
          found = false
          break
        }
      }
      if (found) return i
    }
  }

  return -1
}

Buffer.prototype.includes = function includes (val, byteOffset, encoding) {
  return this.indexOf(val, byteOffset, encoding) !== -1
}

Buffer.prototype.indexOf = function indexOf (val, byteOffset, encoding) {
  return bidirectionalIndexOf(this, val, byteOffset, encoding, true)
}

Buffer.prototype.lastIndexOf = function lastIndexOf (val, byteOffset, encoding) {
  return bidirectionalIndexOf(this, val, byteOffset, encoding, false)
}

function hexWrite (buf, string, offset, length) {
  offset = Number(offset) || 0
  var remaining = buf.length - offset
  if (!length) {
    length = remaining
  } else {
    length = Number(length)
    if (length > remaining) {
      length = remaining
    }
  }

  var strLen = string.length

  if (length > strLen / 2) {
    length = strLen / 2
  }
  for (var i = 0; i < length; ++i) {
    var parsed = parseInt(string.substr(i * 2, 2), 16)
    if (numberIsNaN(parsed)) return i
    buf[offset + i] = parsed
  }
  return i
}

function utf8Write (buf, string, offset, length) {
  return blitBuffer(utf8ToBytes(string, buf.length - offset), buf, offset, length)
}

function asciiWrite (buf, string, offset, length) {
  return blitBuffer(asciiToBytes(string), buf, offset, length)
}

function latin1Write (buf, string, offset, length) {
  return asciiWrite(buf, string, offset, length)
}

function base64Write (buf, string, offset, length) {
  return blitBuffer(base64ToBytes(string), buf, offset, length)
}

function ucs2Write (buf, string, offset, length) {
  return blitBuffer(utf16leToBytes(string, buf.length - offset), buf, offset, length)
}

Buffer.prototype.write = function write (string, offset, length, encoding) {
  // Buffer#write(string)
  if (offset === undefined) {
    encoding = 'utf8'
    length = this.length
    offset = 0
  // Buffer#write(string, encoding)
  } else if (length === undefined && typeof offset === 'string') {
    encoding = offset
    length = this.length
    offset = 0
  // Buffer#write(string, offset[, length][, encoding])
  } else if (isFinite(offset)) {
    offset = offset >>> 0
    if (isFinite(length)) {
      length = length >>> 0
      if (encoding === undefined) encoding = 'utf8'
    } else {
      encoding = length
      length = undefined
    }
  } else {
    throw new Error(
      'Buffer.write(string, encoding, offset[, length]) is no longer supported'
    )
  }

  var remaining = this.length - offset
  if (length === undefined || length > remaining) length = remaining

  if ((string.length > 0 && (length < 0 || offset < 0)) || offset > this.length) {
    throw new RangeError('Attempt to write outside buffer bounds')
  }

  if (!encoding) encoding = 'utf8'

  var loweredCase = false
  for (;;) {
    switch (encoding) {
      case 'hex':
        return hexWrite(this, string, offset, length)

      case 'utf8':
      case 'utf-8':
        return utf8Write(this, string, offset, length)

      case 'ascii':
        return asciiWrite(this, string, offset, length)

      case 'latin1':
      case 'binary':
        return latin1Write(this, string, offset, length)

      case 'base64':
        // Warning: maxLength not taken into account in base64Write
        return base64Write(this, string, offset, length)

      case 'ucs2':
      case 'ucs-2':
      case 'utf16le':
      case 'utf-16le':
        return ucs2Write(this, string, offset, length)

      default:
        if (loweredCase) throw new TypeError('Unknown encoding: ' + encoding)
        encoding = ('' + encoding).toLowerCase()
        loweredCase = true
    }
  }
}

Buffer.prototype.toJSON = function toJSON () {
  return {
    type: 'Buffer',
    data: Array.prototype.slice.call(this._arr || this, 0)
  }
}

function base64Slice (buf, start, end) {
  if (start === 0 && end === buf.length) {
    return base64.fromByteArray(buf)
  } else {
    return base64.fromByteArray(buf.slice(start, end))
  }
}

function utf8Slice (buf, start, end) {
  end = Math.min(buf.length, end)
  var res = []

  var i = start
  while (i < end) {
    var firstByte = buf[i]
    var codePoint = null
    var bytesPerSequence = (firstByte > 0xEF) ? 4
      : (firstByte > 0xDF) ? 3
        : (firstByte > 0xBF) ? 2
          : 1

    if (i + bytesPerSequence <= end) {
      var secondByte, thirdByte, fourthByte, tempCodePoint

      switch (bytesPerSequence) {
        case 1:
          if (firstByte < 0x80) {
            codePoint = firstByte
          }
          break
        case 2:
          secondByte = buf[i + 1]
          if ((secondByte & 0xC0) === 0x80) {
            tempCodePoint = (firstByte & 0x1F) << 0x6 | (secondByte & 0x3F)
            if (tempCodePoint > 0x7F) {
              codePoint = tempCodePoint
            }
          }
          break
        case 3:
          secondByte = buf[i + 1]
          thirdByte = buf[i + 2]
          if ((secondByte & 0xC0) === 0x80 && (thirdByte & 0xC0) === 0x80) {
            tempCodePoint = (firstByte & 0xF) << 0xC | (secondByte & 0x3F) << 0x6 | (thirdByte & 0x3F)
            if (tempCodePoint > 0x7FF && (tempCodePoint < 0xD800 || tempCodePoint > 0xDFFF)) {
              codePoint = tempCodePoint
            }
          }
          break
        case 4:
          secondByte = buf[i + 1]
          thirdByte = buf[i + 2]
          fourthByte = buf[i + 3]
          if ((secondByte & 0xC0) === 0x80 && (thirdByte & 0xC0) === 0x80 && (fourthByte & 0xC0) === 0x80) {
            tempCodePoint = (firstByte & 0xF) << 0x12 | (secondByte & 0x3F) << 0xC | (thirdByte & 0x3F) << 0x6 | (fourthByte & 0x3F)
            if (tempCodePoint > 0xFFFF && tempCodePoint < 0x110000) {
              codePoint = tempCodePoint
            }
          }
      }
    }

    if (codePoint === null) {
      // we did not generate a valid codePoint so insert a
      // replacement char (U+FFFD) and advance only 1 byte
      codePoint = 0xFFFD
      bytesPerSequence = 1
    } else if (codePoint > 0xFFFF) {
      // encode to utf16 (surrogate pair dance)
      codePoint -= 0x10000
      res.push(codePoint >>> 10 & 0x3FF | 0xD800)
      codePoint = 0xDC00 | codePoint & 0x3FF
    }

    res.push(codePoint)
    i += bytesPerSequence
  }

  return decodeCodePointsArray(res)
}

// Based on http://stackoverflow.com/a/22747272/680742, the browser with
// the lowest limit is Chrome, with 0x10000 args.
// We go 1 magnitude less, for safety
var MAX_ARGUMENTS_LENGTH = 0x1000

function decodeCodePointsArray (codePoints) {
  var len = codePoints.length
  if (len <= MAX_ARGUMENTS_LENGTH) {
    return String.fromCharCode.apply(String, codePoints) // avoid extra slice()
  }

  // Decode in chunks to avoid "call stack size exceeded".
  var res = ''
  var i = 0
  while (i < len) {
    res += String.fromCharCode.apply(
      String,
      codePoints.slice(i, i += MAX_ARGUMENTS_LENGTH)
    )
  }
  return res
}

function asciiSlice (buf, start, end) {
  var ret = ''
  end = Math.min(buf.length, end)

  for (var i = start; i < end; ++i) {
    ret += String.fromCharCode(buf[i] & 0x7F)
  }
  return ret
}

function latin1Slice (buf, start, end) {
  var ret = ''
  end = Math.min(buf.length, end)

  for (var i = start; i < end; ++i) {
    ret += String.fromCharCode(buf[i])
  }
  return ret
}

function hexSlice (buf, start, end) {
  var len = buf.length

  if (!start || start < 0) start = 0
  if (!end || end < 0 || end > len) end = len

  var out = ''
  for (var i = start; i < end; ++i) {
    out += toHex(buf[i])
  }
  return out
}

function utf16leSlice (buf, start, end) {
  var bytes = buf.slice(start, end)
  var res = ''
  for (var i = 0; i < bytes.length; i += 2) {
    res += String.fromCharCode(bytes[i] + (bytes[i + 1] * 256))
  }
  return res
}

Buffer.prototype.slice = function slice (start, end) {
  var len = this.length
  start = ~~start
  end = end === undefined ? len : ~~end

  if (start < 0) {
    start += len
    if (start < 0) start = 0
  } else if (start > len) {
    start = len
  }

  if (end < 0) {
    end += len
    if (end < 0) end = 0
  } else if (end > len) {
    end = len
  }

  if (end < start) end = start

  var newBuf = this.subarray(start, end)
  // Return an augmented `Uint8Array` instance
  newBuf.__proto__ = Buffer.prototype
  return newBuf
}

/*
 * Need to make sure that buffer isn't trying to write out of bounds.
 */
function checkOffset (offset, ext, length) {
  if ((offset % 1) !== 0 || offset < 0) throw new RangeError('offset is not uint')
  if (offset + ext > length) throw new RangeError('Trying to access beyond buffer length')
}

Buffer.prototype.readUIntLE = function readUIntLE (offset, byteLength, noAssert) {
  offset = offset >>> 0
  byteLength = byteLength >>> 0
  if (!noAssert) checkOffset(offset, byteLength, this.length)

  var val = this[offset]
  var mul = 1
  var i = 0
  while (++i < byteLength && (mul *= 0x100)) {
    val += this[offset + i] * mul
  }

  return val
}

Buffer.prototype.readUIntBE = function readUIntBE (offset, byteLength, noAssert) {
  offset = offset >>> 0
  byteLength = byteLength >>> 0
  if (!noAssert) {
    checkOffset(offset, byteLength, this.length)
  }

  var val = this[offset + --byteLength]
  var mul = 1
  while (byteLength > 0 && (mul *= 0x100)) {
    val += this[offset + --byteLength] * mul
  }

  return val
}

Buffer.prototype.readUInt8 = function readUInt8 (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 1, this.length)
  return this[offset]
}

Buffer.prototype.readUInt16LE = function readUInt16LE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 2, this.length)
  return this[offset] | (this[offset + 1] << 8)
}

Buffer.prototype.readUInt16BE = function readUInt16BE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 2, this.length)
  return (this[offset] << 8) | this[offset + 1]
}

Buffer.prototype.readUInt32LE = function readUInt32LE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 4, this.length)

  return ((this[offset]) |
      (this[offset + 1] << 8) |
      (this[offset + 2] << 16)) +
      (this[offset + 3] * 0x1000000)
}

Buffer.prototype.readUInt32BE = function readUInt32BE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 4, this.length)

  return (this[offset] * 0x1000000) +
    ((this[offset + 1] << 16) |
    (this[offset + 2] << 8) |
    this[offset + 3])
}

Buffer.prototype.readIntLE = function readIntLE (offset, byteLength, noAssert) {
  offset = offset >>> 0
  byteLength = byteLength >>> 0
  if (!noAssert) checkOffset(offset, byteLength, this.length)

  var val = this[offset]
  var mul = 1
  var i = 0
  while (++i < byteLength && (mul *= 0x100)) {
    val += this[offset + i] * mul
  }
  mul *= 0x80

  if (val >= mul) val -= Math.pow(2, 8 * byteLength)

  return val
}

Buffer.prototype.readIntBE = function readIntBE (offset, byteLength, noAssert) {
  offset = offset >>> 0
  byteLength = byteLength >>> 0
  if (!noAssert) checkOffset(offset, byteLength, this.length)

  var i = byteLength
  var mul = 1
  var val = this[offset + --i]
  while (i > 0 && (mul *= 0x100)) {
    val += this[offset + --i] * mul
  }
  mul *= 0x80

  if (val >= mul) val -= Math.pow(2, 8 * byteLength)

  return val
}

Buffer.prototype.readInt8 = function readInt8 (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 1, this.length)
  if (!(this[offset] & 0x80)) return (this[offset])
  return ((0xff - this[offset] + 1) * -1)
}

Buffer.prototype.readInt16LE = function readInt16LE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 2, this.length)
  var val = this[offset] | (this[offset + 1] << 8)
  return (val & 0x8000) ? val | 0xFFFF0000 : val
}

Buffer.prototype.readInt16BE = function readInt16BE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 2, this.length)
  var val = this[offset + 1] | (this[offset] << 8)
  return (val & 0x8000) ? val | 0xFFFF0000 : val
}

Buffer.prototype.readInt32LE = function readInt32LE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 4, this.length)

  return (this[offset]) |
    (this[offset + 1] << 8) |
    (this[offset + 2] << 16) |
    (this[offset + 3] << 24)
}

Buffer.prototype.readInt32BE = function readInt32BE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 4, this.length)

  return (this[offset] << 24) |
    (this[offset + 1] << 16) |
    (this[offset + 2] << 8) |
    (this[offset + 3])
}

Buffer.prototype.readFloatLE = function readFloatLE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 4, this.length)
  return ieee754.read(this, offset, true, 23, 4)
}

Buffer.prototype.readFloatBE = function readFloatBE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 4, this.length)
  return ieee754.read(this, offset, false, 23, 4)
}

Buffer.prototype.readDoubleLE = function readDoubleLE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 8, this.length)
  return ieee754.read(this, offset, true, 52, 8)
}

Buffer.prototype.readDoubleBE = function readDoubleBE (offset, noAssert) {
  offset = offset >>> 0
  if (!noAssert) checkOffset(offset, 8, this.length)
  return ieee754.read(this, offset, false, 52, 8)
}

function checkInt (buf, value, offset, ext, max, min) {
  if (!Buffer.isBuffer(buf)) throw new TypeError('"buffer" argument must be a Buffer instance')
  if (value > max || value < min) throw new RangeError('"value" argument is out of bounds')
  if (offset + ext > buf.length) throw new RangeError('Index out of range')
}

Buffer.prototype.writeUIntLE = function writeUIntLE (value, offset, byteLength, noAssert) {
  value = +value
  offset = offset >>> 0
  byteLength = byteLength >>> 0
  if (!noAssert) {
    var maxBytes = Math.pow(2, 8 * byteLength) - 1
    checkInt(this, value, offset, byteLength, maxBytes, 0)
  }

  var mul = 1
  var i = 0
  this[offset] = value & 0xFF
  while (++i < byteLength && (mul *= 0x100)) {
    this[offset + i] = (value / mul) & 0xFF
  }

  return offset + byteLength
}

Buffer.prototype.writeUIntBE = function writeUIntBE (value, offset, byteLength, noAssert) {
  value = +value
  offset = offset >>> 0
  byteLength = byteLength >>> 0
  if (!noAssert) {
    var maxBytes = Math.pow(2, 8 * byteLength) - 1
    checkInt(this, value, offset, byteLength, maxBytes, 0)
  }

  var i = byteLength - 1
  var mul = 1
  this[offset + i] = value & 0xFF
  while (--i >= 0 && (mul *= 0x100)) {
    this[offset + i] = (value / mul) & 0xFF
  }

  return offset + byteLength
}

Buffer.prototype.writeUInt8 = function writeUInt8 (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 1, 0xff, 0)
  this[offset] = (value & 0xff)
  return offset + 1
}

Buffer.prototype.writeUInt16LE = function writeUInt16LE (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 2, 0xffff, 0)
  this[offset] = (value & 0xff)
  this[offset + 1] = (value >>> 8)
  return offset + 2
}

Buffer.prototype.writeUInt16BE = function writeUInt16BE (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 2, 0xffff, 0)
  this[offset] = (value >>> 8)
  this[offset + 1] = (value & 0xff)
  return offset + 2
}

Buffer.prototype.writeUInt32LE = function writeUInt32LE (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 4, 0xffffffff, 0)
  this[offset + 3] = (value >>> 24)
  this[offset + 2] = (value >>> 16)
  this[offset + 1] = (value >>> 8)
  this[offset] = (value & 0xff)
  return offset + 4
}

Buffer.prototype.writeUInt32BE = function writeUInt32BE (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 4, 0xffffffff, 0)
  this[offset] = (value >>> 24)
  this[offset + 1] = (value >>> 16)
  this[offset + 2] = (value >>> 8)
  this[offset + 3] = (value & 0xff)
  return offset + 4
}

Buffer.prototype.writeIntLE = function writeIntLE (value, offset, byteLength, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) {
    var limit = Math.pow(2, (8 * byteLength) - 1)

    checkInt(this, value, offset, byteLength, limit - 1, -limit)
  }

  var i = 0
  var mul = 1
  var sub = 0
  this[offset] = value & 0xFF
  while (++i < byteLength && (mul *= 0x100)) {
    if (value < 0 && sub === 0 && this[offset + i - 1] !== 0) {
      sub = 1
    }
    this[offset + i] = ((value / mul) >> 0) - sub & 0xFF
  }

  return offset + byteLength
}

Buffer.prototype.writeIntBE = function writeIntBE (value, offset, byteLength, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) {
    var limit = Math.pow(2, (8 * byteLength) - 1)

    checkInt(this, value, offset, byteLength, limit - 1, -limit)
  }

  var i = byteLength - 1
  var mul = 1
  var sub = 0
  this[offset + i] = value & 0xFF
  while (--i >= 0 && (mul *= 0x100)) {
    if (value < 0 && sub === 0 && this[offset + i + 1] !== 0) {
      sub = 1
    }
    this[offset + i] = ((value / mul) >> 0) - sub & 0xFF
  }

  return offset + byteLength
}

Buffer.prototype.writeInt8 = function writeInt8 (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 1, 0x7f, -0x80)
  if (value < 0) value = 0xff + value + 1
  this[offset] = (value & 0xff)
  return offset + 1
}

Buffer.prototype.writeInt16LE = function writeInt16LE (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 2, 0x7fff, -0x8000)
  this[offset] = (value & 0xff)
  this[offset + 1] = (value >>> 8)
  return offset + 2
}

Buffer.prototype.writeInt16BE = function writeInt16BE (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 2, 0x7fff, -0x8000)
  this[offset] = (value >>> 8)
  this[offset + 1] = (value & 0xff)
  return offset + 2
}

Buffer.prototype.writeInt32LE = function writeInt32LE (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 4, 0x7fffffff, -0x80000000)
  this[offset] = (value & 0xff)
  this[offset + 1] = (value >>> 8)
  this[offset + 2] = (value >>> 16)
  this[offset + 3] = (value >>> 24)
  return offset + 4
}

Buffer.prototype.writeInt32BE = function writeInt32BE (value, offset, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) checkInt(this, value, offset, 4, 0x7fffffff, -0x80000000)
  if (value < 0) value = 0xffffffff + value + 1
  this[offset] = (value >>> 24)
  this[offset + 1] = (value >>> 16)
  this[offset + 2] = (value >>> 8)
  this[offset + 3] = (value & 0xff)
  return offset + 4
}

function checkIEEE754 (buf, value, offset, ext, max, min) {
  if (offset + ext > buf.length) throw new RangeError('Index out of range')
  if (offset < 0) throw new RangeError('Index out of range')
}

function writeFloat (buf, value, offset, littleEndian, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) {
    checkIEEE754(buf, value, offset, 4, 3.4028234663852886e+38, -3.4028234663852886e+38)
  }
  ieee754.write(buf, value, offset, littleEndian, 23, 4)
  return offset + 4
}

Buffer.prototype.writeFloatLE = function writeFloatLE (value, offset, noAssert) {
  return writeFloat(this, value, offset, true, noAssert)
}

Buffer.prototype.writeFloatBE = function writeFloatBE (value, offset, noAssert) {
  return writeFloat(this, value, offset, false, noAssert)
}

function writeDouble (buf, value, offset, littleEndian, noAssert) {
  value = +value
  offset = offset >>> 0
  if (!noAssert) {
    checkIEEE754(buf, value, offset, 8, 1.7976931348623157E+308, -1.7976931348623157E+308)
  }
  ieee754.write(buf, value, offset, littleEndian, 52, 8)
  return offset + 8
}

Buffer.prototype.writeDoubleLE = function writeDoubleLE (value, offset, noAssert) {
  return writeDouble(this, value, offset, true, noAssert)
}

Buffer.prototype.writeDoubleBE = function writeDoubleBE (value, offset, noAssert) {
  return writeDouble(this, value, offset, false, noAssert)
}

// copy(targetBuffer, targetStart=0, sourceStart=0, sourceEnd=buffer.length)
Buffer.prototype.copy = function copy (target, targetStart, start, end) {
  if (!Buffer.isBuffer(target)) throw new TypeError('argument should be a Buffer')
  if (!start) start = 0
  if (!end && end !== 0) end = this.length
  if (targetStart >= target.length) targetStart = target.length
  if (!targetStart) targetStart = 0
  if (end > 0 && end < start) end = start

  // Copy 0 bytes; we're done
  if (end === start) return 0
  if (target.length === 0 || this.length === 0) return 0

  // Fatal error conditions
  if (targetStart < 0) {
    throw new RangeError('targetStart out of bounds')
  }
  if (start < 0 || start >= this.length) throw new RangeError('Index out of range')
  if (end < 0) throw new RangeError('sourceEnd out of bounds')

  // Are we oob?
  if (end > this.length) end = this.length
  if (target.length - targetStart < end - start) {
    end = target.length - targetStart + start
  }

  var len = end - start

  if (this === target && typeof Uint8Array.prototype.copyWithin === 'function') {
    // Use built-in when available, missing from IE11
    this.copyWithin(targetStart, start, end)
  } else if (this === target && start < targetStart && targetStart < end) {
    // descending copy from end
    for (var i = len - 1; i >= 0; --i) {
      target[i + targetStart] = this[i + start]
    }
  } else {
    Uint8Array.prototype.set.call(
      target,
      this.subarray(start, end),
      targetStart
    )
  }

  return len
}

// Usage:
//    buffer.fill(number[, offset[, end]])
//    buffer.fill(buffer[, offset[, end]])
//    buffer.fill(string[, offset[, end]][, encoding])
Buffer.prototype.fill = function fill (val, start, end, encoding) {
  // Handle string cases:
  if (typeof val === 'string') {
    if (typeof start === 'string') {
      encoding = start
      start = 0
      end = this.length
    } else if (typeof end === 'string') {
      encoding = end
      end = this.length
    }
    if (encoding !== undefined && typeof encoding !== 'string') {
      throw new TypeError('encoding must be a string')
    }
    if (typeof encoding === 'string' && !Buffer.isEncoding(encoding)) {
      throw new TypeError('Unknown encoding: ' + encoding)
    }
    if (val.length === 1) {
      var code = val.charCodeAt(0)
      if ((encoding === 'utf8' && code < 128) ||
          encoding === 'latin1') {
        // Fast path: If `val` fits into a single byte, use that numeric value.
        val = code
      }
    }
  } else if (typeof val === 'number') {
    val = val & 255
  }

  // Invalid ranges are not set to a default, so can range check early.
  if (start < 0 || this.length < start || this.length < end) {
    throw new RangeError('Out of range index')
  }

  if (end <= start) {
    return this
  }

  start = start >>> 0
  end = end === undefined ? this.length : end >>> 0

  if (!val) val = 0

  var i
  if (typeof val === 'number') {
    for (i = start; i < end; ++i) {
      this[i] = val
    }
  } else {
    var bytes = Buffer.isBuffer(val)
      ? val
      : Buffer.from(val, encoding)
    var len = bytes.length
    if (len === 0) {
      throw new TypeError('The value "' + val +
        '" is invalid for argument "value"')
    }
    for (i = 0; i < end - start; ++i) {
      this[i + start] = bytes[i % len]
    }
  }

  return this
}

// HELPER FUNCTIONS
// ================

var INVALID_BASE64_RE = /[^+/0-9A-Za-z-_]/g

function base64clean (str) {
  // Node takes equal signs as end of the Base64 encoding
  str = str.split('=')[0]
  // Node strips out invalid characters like \n and \t from the string, base64-js does not
  str = str.trim().replace(INVALID_BASE64_RE, '')
  // Node converts strings with length < 2 to ''
  if (str.length < 2) return ''
  // Node allows for non-padded base64 strings (missing trailing ===), base64-js does not
  while (str.length % 4 !== 0) {
    str = str + '='
  }
  return str
}

function toHex (n) {
  if (n < 16) return '0' + n.toString(16)
  return n.toString(16)
}

function utf8ToBytes (string, units) {
  units = units || Infinity
  var codePoint
  var length = string.length
  var leadSurrogate = null
  var bytes = []

  for (var i = 0; i < length; ++i) {
    codePoint = string.charCodeAt(i)

    // is surrogate component
    if (codePoint > 0xD7FF && codePoint < 0xE000) {
      // last char was a lead
      if (!leadSurrogate) {
        // no lead yet
        if (codePoint > 0xDBFF) {
          // unexpected trail
          if ((units -= 3) > -1) bytes.push(0xEF, 0xBF, 0xBD)
          continue
        } else if (i + 1 === length) {
          // unpaired lead
          if ((units -= 3) > -1) bytes.push(0xEF, 0xBF, 0xBD)
          continue
        }

        // valid lead
        leadSurrogate = codePoint

        continue
      }

      // 2 leads in a row
      if (codePoint < 0xDC00) {
        if ((units -= 3) > -1) bytes.push(0xEF, 0xBF, 0xBD)
        leadSurrogate = codePoint
        continue
      }

      // valid surrogate pair
      codePoint = (leadSurrogate - 0xD800 << 10 | codePoint - 0xDC00) + 0x10000
    } else if (leadSurrogate) {
      // valid bmp char, but last char was a lead
      if ((units -= 3) > -1) bytes.push(0xEF, 0xBF, 0xBD)
    }

    leadSurrogate = null

    // encode utf8
    if (codePoint < 0x80) {
      if ((units -= 1) < 0) break
      bytes.push(codePoint)
    } else if (codePoint < 0x800) {
      if ((units -= 2) < 0) break
      bytes.push(
        codePoint >> 0x6 | 0xC0,
        codePoint & 0x3F | 0x80
      )
    } else if (codePoint < 0x10000) {
      if ((units -= 3) < 0) break
      bytes.push(
        codePoint >> 0xC | 0xE0,
        codePoint >> 0x6 & 0x3F | 0x80,
        codePoint & 0x3F | 0x80
      )
    } else if (codePoint < 0x110000) {
      if ((units -= 4) < 0) break
      bytes.push(
        codePoint >> 0x12 | 0xF0,
        codePoint >> 0xC & 0x3F | 0x80,
        codePoint >> 0x6 & 0x3F | 0x80,
        codePoint & 0x3F | 0x80
      )
    } else {
      throw new Error('Invalid code point')
    }
  }

  return bytes
}

function asciiToBytes (str) {
  var byteArray = []
  for (var i = 0; i < str.length; ++i) {
    // Node's code seems to be doing this and not & 0x7F..
    byteArray.push(str.charCodeAt(i) & 0xFF)
  }
  return byteArray
}

function utf16leToBytes (str, units) {
  var c, hi, lo
  var byteArray = []
  for (var i = 0; i < str.length; ++i) {
    if ((units -= 2) < 0) break

    c = str.charCodeAt(i)
    hi = c >> 8
    lo = c % 256
    byteArray.push(lo)
    byteArray.push(hi)
  }

  return byteArray
}

function base64ToBytes (str) {
  return base64.toByteArray(base64clean(str))
}

function blitBuffer (src, dst, offset, length) {
  for (var i = 0; i < length; ++i) {
    if ((i + offset >= dst.length) || (i >= src.length)) break
    dst[i + offset] = src[i]
  }
  return i
}

// ArrayBuffer or Uint8Array objects from other contexts (i.e. iframes) do not pass
// the `instanceof` check but they should be treated as of that type.
// See: https://github.com/feross/buffer/issues/166
function isInstance (obj, type) {
  return obj instanceof type ||
    (obj != null && obj.constructor != null && obj.constructor.name != null &&
      obj.constructor.name === type.name)
}
function numberIsNaN (obj) {
  // For IE11 support
  return obj !== obj // eslint-disable-line no-self-compare
}

}).call(this)}).call(this,require("buffer").Buffer)
},{"base64-js":2,"buffer":3,"ieee754":5}],4:[function(require,module,exports){
// Copyright Joyent, Inc. and other Node contributors.
//
// Permission is hereby granted, free of charge, to any person obtaining a
// copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to permit
// persons to whom the Software is furnished to do so, subject to the
// following conditions:
//
// The above copyright notice and this permission notice shall be included
// in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
// OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN
// NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
// DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
// OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
// USE OR OTHER DEALINGS IN THE SOFTWARE.

'use strict';

var R = typeof Reflect === 'object' ? Reflect : null
var ReflectApply = R && typeof R.apply === 'function'
  ? R.apply
  : function ReflectApply(target, receiver, args) {
    return Function.prototype.apply.call(target, receiver, args);
  }

var ReflectOwnKeys
if (R && typeof R.ownKeys === 'function') {
  ReflectOwnKeys = R.ownKeys
} else if (Object.getOwnPropertySymbols) {
  ReflectOwnKeys = function ReflectOwnKeys(target) {
    return Object.getOwnPropertyNames(target)
      .concat(Object.getOwnPropertySymbols(target));
  };
} else {
  ReflectOwnKeys = function ReflectOwnKeys(target) {
    return Object.getOwnPropertyNames(target);
  };
}

function ProcessEmitWarning(warning) {
  if (console && console.warn) console.warn(warning);
}

var NumberIsNaN = Number.isNaN || function NumberIsNaN(value) {
  return value !== value;
}

function EventEmitter() {
  EventEmitter.init.call(this);
}
module.exports = EventEmitter;
module.exports.once = once;

// Backwards-compat with node 0.10.x
EventEmitter.EventEmitter = EventEmitter;

EventEmitter.prototype._events = undefined;
EventEmitter.prototype._eventsCount = 0;
EventEmitter.prototype._maxListeners = undefined;

// By default EventEmitters will print a warning if more than 10 listeners are
// added to it. This is a useful default which helps finding memory leaks.
var defaultMaxListeners = 10;

function checkListener(listener) {
  if (typeof listener !== 'function') {
    throw new TypeError('The "listener" argument must be of type Function. Received type ' + typeof listener);
  }
}

Object.defineProperty(EventEmitter, 'defaultMaxListeners', {
  enumerable: true,
  get: function() {
    return defaultMaxListeners;
  },
  set: function(arg) {
    if (typeof arg !== 'number' || arg < 0 || NumberIsNaN(arg)) {
      throw new RangeError('The value of "defaultMaxListeners" is out of range. It must be a non-negative number. Received ' + arg + '.');
    }
    defaultMaxListeners = arg;
  }
});

EventEmitter.init = function() {

  if (this._events === undefined ||
      this._events === Object.getPrototypeOf(this)._events) {
    this._events = Object.create(null);
    this._eventsCount = 0;
  }

  this._maxListeners = this._maxListeners || undefined;
};

// Obviously not all Emitters should be limited to 10. This function allows
// that to be increased. Set to zero for unlimited.
EventEmitter.prototype.setMaxListeners = function setMaxListeners(n) {
  if (typeof n !== 'number' || n < 0 || NumberIsNaN(n)) {
    throw new RangeError('The value of "n" is out of range. It must be a non-negative number. Received ' + n + '.');
  }
  this._maxListeners = n;
  return this;
};

function _getMaxListeners(that) {
  if (that._maxListeners === undefined)
    return EventEmitter.defaultMaxListeners;
  return that._maxListeners;
}

EventEmitter.prototype.getMaxListeners = function getMaxListeners() {
  return _getMaxListeners(this);
};

EventEmitter.prototype.emit = function emit(type) {
  var args = [];
  for (var i = 1; i < arguments.length; i++) args.push(arguments[i]);
  var doError = (type === 'error');

  var events = this._events;
  if (events !== undefined)
    doError = (doError && events.error === undefined);
  else if (!doError)
    return false;

  // If there is no 'error' event listener then throw.
  if (doError) {
    var er;
    if (args.length > 0)
      er = args[0];
    if (er instanceof Error) {
      // Note: The comments on the `throw` lines are intentional, they show
      // up in Node's output if this results in an unhandled exception.
      throw er; // Unhandled 'error' event
    }
    // At least give some kind of context to the user
    var err = new Error('Unhandled error.' + (er ? ' (' + er.message + ')' : ''));
    err.context = er;
    throw err; // Unhandled 'error' event
  }

  var handler = events[type];

  if (handler === undefined)
    return false;

  if (typeof handler === 'function') {
    ReflectApply(handler, this, args);
  } else {
    var len = handler.length;
    var listeners = arrayClone(handler, len);
    for (var i = 0; i < len; ++i)
      ReflectApply(listeners[i], this, args);
  }

  return true;
};

function _addListener(target, type, listener, prepend) {
  var m;
  var events;
  var existing;

  checkListener(listener);

  events = target._events;
  if (events === undefined) {
    events = target._events = Object.create(null);
    target._eventsCount = 0;
  } else {
    // To avoid recursion in the case that type === "newListener"! Before
    // adding it to the listeners, first emit "newListener".
    if (events.newListener !== undefined) {
      target.emit('newListener', type,
                  listener.listener ? listener.listener : listener);

      // Re-assign `events` because a newListener handler could have caused the
      // this._events to be assigned to a new object
      events = target._events;
    }
    existing = events[type];
  }

  if (existing === undefined) {
    // Optimize the case of one listener. Don't need the extra array object.
    existing = events[type] = listener;
    ++target._eventsCount;
  } else {
    if (typeof existing === 'function') {
      // Adding the second element, need to change to array.
      existing = events[type] =
        prepend ? [listener, existing] : [existing, listener];
      // If we've already got an array, just append.
    } else if (prepend) {
      existing.unshift(listener);
    } else {
      existing.push(listener);
    }

    // Check for listener leak
    m = _getMaxListeners(target);
    if (m > 0 && existing.length > m && !existing.warned) {
      existing.warned = true;
      // No error code for this since it is a Warning
      // eslint-disable-next-line no-restricted-syntax
      var w = new Error('Possible EventEmitter memory leak detected. ' +
                          existing.length + ' ' + String(type) + ' listeners ' +
                          'added. Use emitter.setMaxListeners() to ' +
                          'increase limit');
      w.name = 'MaxListenersExceededWarning';
      w.emitter = target;
      w.type = type;
      w.count = existing.length;
      ProcessEmitWarning(w);
    }
  }

  return target;
}

EventEmitter.prototype.addListener = function addListener(type, listener) {
  return _addListener(this, type, listener, false);
};

EventEmitter.prototype.on = EventEmitter.prototype.addListener;

EventEmitter.prototype.prependListener =
    function prependListener(type, listener) {
      return _addListener(this, type, listener, true);
    };

function onceWrapper() {
  if (!this.fired) {
    this.target.removeListener(this.type, this.wrapFn);
    this.fired = true;
    if (arguments.length === 0)
      return this.listener.call(this.target);
    return this.listener.apply(this.target, arguments);
  }
}

function _onceWrap(target, type, listener) {
  var state = { fired: false, wrapFn: undefined, target: target, type: type, listener: listener };
  var wrapped = onceWrapper.bind(state);
  wrapped.listener = listener;
  state.wrapFn = wrapped;
  return wrapped;
}

EventEmitter.prototype.once = function once(type, listener) {
  checkListener(listener);
  this.on(type, _onceWrap(this, type, listener));
  return this;
};

EventEmitter.prototype.prependOnceListener =
    function prependOnceListener(type, listener) {
      checkListener(listener);
      this.prependListener(type, _onceWrap(this, type, listener));
      return this;
    };

// Emits a 'removeListener' event if and only if the listener was removed.
EventEmitter.prototype.removeListener =
    function removeListener(type, listener) {
      var list, events, position, i, originalListener;

      checkListener(listener);

      events = this._events;
      if (events === undefined)
        return this;

      list = events[type];
      if (list === undefined)
        return this;

      if (list === listener || list.listener === listener) {
        if (--this._eventsCount === 0)
          this._events = Object.create(null);
        else {
          delete events[type];
          if (events.removeListener)
            this.emit('removeListener', type, list.listener || listener);
        }
      } else if (typeof list !== 'function') {
        position = -1;

        for (i = list.length - 1; i >= 0; i--) {
          if (list[i] === listener || list[i].listener === listener) {
            originalListener = list[i].listener;
            position = i;
            break;
          }
        }

        if (position < 0)
          return this;

        if (position === 0)
          list.shift();
        else {
          spliceOne(list, position);
        }

        if (list.length === 1)
          events[type] = list[0];

        if (events.removeListener !== undefined)
          this.emit('removeListener', type, originalListener || listener);
      }

      return this;
    };

EventEmitter.prototype.off = EventEmitter.prototype.removeListener;

EventEmitter.prototype.removeAllListeners =
    function removeAllListeners(type) {
      var listeners, events, i;

      events = this._events;
      if (events === undefined)
        return this;

      // not listening for removeListener, no need to emit
      if (events.removeListener === undefined) {
        if (arguments.length === 0) {
          this._events = Object.create(null);
          this._eventsCount = 0;
        } else if (events[type] !== undefined) {
          if (--this._eventsCount === 0)
            this._events = Object.create(null);
          else
            delete events[type];
        }
        return this;
      }

      // emit removeListener for all listeners on all events
      if (arguments.length === 0) {
        var keys = Object.keys(events);
        var key;
        for (i = 0; i < keys.length; ++i) {
          key = keys[i];
          if (key === 'removeListener') continue;
          this.removeAllListeners(key);
        }
        this.removeAllListeners('removeListener');
        this._events = Object.create(null);
        this._eventsCount = 0;
        return this;
      }

      listeners = events[type];

      if (typeof listeners === 'function') {
        this.removeListener(type, listeners);
      } else if (listeners !== undefined) {
        // LIFO order
        for (i = listeners.length - 1; i >= 0; i--) {
          this.removeListener(type, listeners[i]);
        }
      }

      return this;
    };

function _listeners(target, type, unwrap) {
  var events = target._events;

  if (events === undefined)
    return [];

  var evlistener = events[type];
  if (evlistener === undefined)
    return [];

  if (typeof evlistener === 'function')
    return unwrap ? [evlistener.listener || evlistener] : [evlistener];

  return unwrap ?
    unwrapListeners(evlistener) : arrayClone(evlistener, evlistener.length);
}

EventEmitter.prototype.listeners = function listeners(type) {
  return _listeners(this, type, true);
};

EventEmitter.prototype.rawListeners = function rawListeners(type) {
  return _listeners(this, type, false);
};

EventEmitter.listenerCount = function(emitter, type) {
  if (typeof emitter.listenerCount === 'function') {
    return emitter.listenerCount(type);
  } else {
    return listenerCount.call(emitter, type);
  }
};

EventEmitter.prototype.listenerCount = listenerCount;
function listenerCount(type) {
  var events = this._events;

  if (events !== undefined) {
    var evlistener = events[type];

    if (typeof evlistener === 'function') {
      return 1;
    } else if (evlistener !== undefined) {
      return evlistener.length;
    }
  }

  return 0;
}

EventEmitter.prototype.eventNames = function eventNames() {
  return this._eventsCount > 0 ? ReflectOwnKeys(this._events) : [];
};

function arrayClone(arr, n) {
  var copy = new Array(n);
  for (var i = 0; i < n; ++i)
    copy[i] = arr[i];
  return copy;
}

function spliceOne(list, index) {
  for (; index + 1 < list.length; index++)
    list[index] = list[index + 1];
  list.pop();
}

function unwrapListeners(arr) {
  var ret = new Array(arr.length);
  for (var i = 0; i < ret.length; ++i) {
    ret[i] = arr[i].listener || arr[i];
  }
  return ret;
}

function once(emitter, name) {
  return new Promise(function (resolve, reject) {
    function errorListener(err) {
      emitter.removeListener(name, resolver);
      reject(err);
    }

    function resolver() {
      if (typeof emitter.removeListener === 'function') {
        emitter.removeListener('error', errorListener);
      }
      resolve([].slice.call(arguments));
    };

    eventTargetAgnosticAddListener(emitter, name, resolver, { once: true });
    if (name !== 'error') {
      addErrorHandlerIfEventEmitter(emitter, errorListener, { once: true });
    }
  });
}

function addErrorHandlerIfEventEmitter(emitter, handler, flags) {
  if (typeof emitter.on === 'function') {
    eventTargetAgnosticAddListener(emitter, 'error', handler, flags);
  }
}

function eventTargetAgnosticAddListener(emitter, name, listener, flags) {
  if (typeof emitter.on === 'function') {
    if (flags.once) {
      emitter.once(name, listener);
    } else {
      emitter.on(name, listener);
    }
  } else if (typeof emitter.addEventListener === 'function') {
    // EventTarget does not have `error` event semantics like Node
    // EventEmitters, we do not listen for `error` events here.
    emitter.addEventListener(name, function wrapListener(arg) {
      // IE does not have builtin `{ once: true }` support so we
      // have to do it manually.
      if (flags.once) {
        emitter.removeEventListener(name, wrapListener);
      }
      listener(arg);
    });
  } else {
    throw new TypeError('The "emitter" argument must be of type EventEmitter. Received type ' + typeof emitter);
  }
}

},{}],5:[function(require,module,exports){
/*! ieee754. BSD-3-Clause License. Feross Aboukhadijeh <https://feross.org/opensource> */
exports.read = function (buffer, offset, isLE, mLen, nBytes) {
  var e, m
  var eLen = (nBytes * 8) - mLen - 1
  var eMax = (1 << eLen) - 1
  var eBias = eMax >> 1
  var nBits = -7
  var i = isLE ? (nBytes - 1) : 0
  var d = isLE ? -1 : 1
  var s = buffer[offset + i]

  i += d

  e = s & ((1 << (-nBits)) - 1)
  s >>= (-nBits)
  nBits += eLen
  for (; nBits > 0; e = (e * 256) + buffer[offset + i], i += d, nBits -= 8) {}

  m = e & ((1 << (-nBits)) - 1)
  e >>= (-nBits)
  nBits += mLen
  for (; nBits > 0; m = (m * 256) + buffer[offset + i], i += d, nBits -= 8) {}

  if (e === 0) {
    e = 1 - eBias
  } else if (e === eMax) {
    return m ? NaN : ((s ? -1 : 1) * Infinity)
  } else {
    m = m + Math.pow(2, mLen)
    e = e - eBias
  }
  return (s ? -1 : 1) * m * Math.pow(2, e - mLen)
}

exports.write = function (buffer, value, offset, isLE, mLen, nBytes) {
  var e, m, c
  var eLen = (nBytes * 8) - mLen - 1
  var eMax = (1 << eLen) - 1
  var eBias = eMax >> 1
  var rt = (mLen === 23 ? Math.pow(2, -24) - Math.pow(2, -77) : 0)
  var i = isLE ? 0 : (nBytes - 1)
  var d = isLE ? 1 : -1
  var s = value < 0 || (value === 0 && 1 / value < 0) ? 1 : 0

  value = Math.abs(value)

  if (isNaN(value) || value === Infinity) {
    m = isNaN(value) ? 1 : 0
    e = eMax
  } else {
    e = Math.floor(Math.log(value) / Math.LN2)
    if (value * (c = Math.pow(2, -e)) < 1) {
      e--
      c *= 2
    }
    if (e + eBias >= 1) {
      value += rt / c
    } else {
      value += rt * Math.pow(2, 1 - eBias)
    }
    if (value * c >= 2) {
      e++
      c /= 2
    }

    if (e + eBias >= eMax) {
      m = 0
      e = eMax
    } else if (e + eBias >= 1) {
      m = ((value * c) - 1) * Math.pow(2, mLen)
      e = e + eBias
    } else {
      m = value * Math.pow(2, eBias - 1) * Math.pow(2, mLen)
      e = 0
    }
  }

  for (; mLen >= 8; buffer[offset + i] = m & 0xff, i += d, m /= 256, mLen -= 8) {}

  e = (e << mLen) | m
  eLen += mLen
  for (; eLen > 0; buffer[offset + i] = e & 0xff, i += d, e /= 256, eLen -= 8) {}

  buffer[offset + i - d] |= s * 128
}

},{}],6:[function(require,module,exports){
/**
 * @module buffer-reader
 * 
 * @description
 * Exports a class {@link BufferReader}, used internally to handle 
 * access when an input buffer is passed. This provides a consistent
 * interface between reading from files and buffers, so that in-memory
 * files can be handled efficiently.
 */

/**
 * A class that allows a reader to access file through the file system.
 * This can be used as an alternative to the 
 * [FileReader]{@link module:file-reader~FileReader} which
 * reads direct from an opened file descriptor. 
 */
class BufferReader {

  constructor(buffer) {
    this._buffer = buffer;
  }

  open() {
    return Promise.resolve();
  }

  close() {
    return Promise.resolve();
  }

  read(buffer, offset, length, position) {
    this._buffer.copy(buffer, offset, position, position + length);
    return Promise.resolve(buffer);
  }

  buffer() {
    return this._buffer;
  }

  static isBufferReader(instance) {
    return instance instanceof BufferReader;
  }
}

module.exports = BufferReader;

},{}],7:[function(require,module,exports){
/**
 * @module document
 * 
 * @description
 * Implements the main document returned when a Word file has been extracted. This exposes
 * methods that allow the body, annotations, headers, footnotes, and endnotes, to be 
 * read and used.
 * 
 * @author
 * Stuart Watt <stuart@morungos.com>
 */

const { filter } = require('./filters');

/**
 * @class
 * Returned from all extractors, this class provides accessors to read the different
 * parts of a Word document. This also allows some options to be passed to the accessors,
 * so you can control some character conversion and filtering, as described in the methods
 * below.
 */
class Document {
  
  constructor() {
    this._body = "";
    this._footnotes = "";
    this._endnotes = "";
    this._headers = "";
    this._footers = "";
    this._annotations = "";
    this._textboxes = "";
    this._headerTextboxes = "";
  }

  /**
   * Accessor to read the main body part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file body
   */
  getBody(options) {
    options = options || {};
    const value = this._body;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the footnotes part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file footnotes
   */
  getFootnotes(options) {
    options = options || {};
    const value = this._footnotes;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the endnotes part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file endnotes
   */
  getEndnotes(options) {
    options = options || {};
    const value = this._endnotes;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the headers part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @param {boolean} options.includeFooters - if true (the default), returns headers and footers 
   *   as a single string
   * @returns a string, containing the Word file headers
   */
  getHeaders(options) {
    options = options || {};
    const value = this._headers + ((options.includeFooters == false) ? "" : this._footers);
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the footers part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file footers
   */
  getFooters(options) {
    options = options || {};
    const value = this._footers;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the annotations part of a Word file
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @returns a string, containing the Word file annotations
   */
  getAnnotations(options) {
    options = options || {};
    const value = this._annotations;
    return (options.filterUnicode == false) ? value : filter(value);
  }

  /**
   * Accessor to read the textboxes from a Word file. The text box content is aggregated as a 
   * single long string. When both the body and header content exists, they will be separated
   * by a newline.
   * @param {Object} options - options for body data
   * @param {boolean} options.filterUnicode - if true (the default), converts common Unicode quotes
   *   to standard ASCII characters 
   * @param {boolean} options.includeHeadersAndFooters - if true (the default), includes text box
   *   content in headers and footers
   * @param {boolean} options.includeBody - if true (the default), includes text box
   *   content in the document body
   * @returns a string, containing the Word file text box content
   */
  getTextboxes(options) {
    options = options || {};
    const segments = [];
    if (options.includeBody != false) 
      segments.push(this._textboxes);
    if (options.includeHeadersAndFooters != false)
      segments.push(this._headerTextboxes);
    const value = segments.join("\n");
    return (options.filterUnicode == false) ? value : filter(value);
  }
}


module.exports = Document;

},{"./filters":8}],8:[function(require,module,exports){
/**
 * @module filters
 * 
 * @description
 * Exports several functions that implement various methods for translating
 * characters into Unicode, and cleaning up some of the remaining residues from
 * Word's odd internal marker character usage.
 */

/**
 * A replacement table, that maps Word control characters to either NULL, for
 * deletion, or to another more acceptable character ina Unicode world, such 
 * as a newline.
 */
const replaceTable = [];
replaceTable[0x0002] = '\x00';
replaceTable[0x0005] = '\x00';
replaceTable[0x0007] = "\t";
replaceTable[0x0008] = '\x00';
replaceTable[0x000A] = "\n";
replaceTable[0x000B] = "\n";
replaceTable[0x000C] = "\n";
replaceTable[0x000D] = "\n";
replaceTable[0x001E] = "\u2011";

/**
 * @constant
 * Maps between Windows character codes, especially between 0x80 and 0x9f,
 * into official Unicode code points. This smooths over the differences
 * between UCS-2 and 8-bit code runs in Word, by allowing us to work
 * entirely within Unicode later on.
 */
const binaryToUnicodeTable = [];
binaryToUnicodeTable[0x0082] = "\u201a";
binaryToUnicodeTable[0x0083] = "\u0192";
binaryToUnicodeTable[0x0084] = "\u201e";
binaryToUnicodeTable[0x0085] = "\u2026";
binaryToUnicodeTable[0x0086] = "\u2020";
binaryToUnicodeTable[0x0087] = "\u2021";
binaryToUnicodeTable[0x0088] = "\u02C6";
binaryToUnicodeTable[0x0089] = "\u2030";
binaryToUnicodeTable[0x008a] = "\u0160";
binaryToUnicodeTable[0x008b] = "\u2039";
binaryToUnicodeTable[0x008c] = "\u0152";
binaryToUnicodeTable[0x008e] = "\u017D";
binaryToUnicodeTable[0x0091] = "\u2018";
binaryToUnicodeTable[0x0092] = "\u2019";
binaryToUnicodeTable[0x0093] = "\u201C";
binaryToUnicodeTable[0x0094] = "\u201D";
binaryToUnicodeTable[0x0095] = "\u2022";
binaryToUnicodeTable[0x0096] = "\u2013";
binaryToUnicodeTable[0x0097] = "\u2014";
binaryToUnicodeTable[0x0098] = "\u02DC";
binaryToUnicodeTable[0x0099] = "\u2122";
binaryToUnicodeTable[0x009a] = "\u0161";
binaryToUnicodeTable[0x009b] = "\u203A";
binaryToUnicodeTable[0x009c] = "\u0153";
binaryToUnicodeTable[0x009e] = "\u017E";
binaryToUnicodeTable[0x009f] = "\u0178";

/**
 * Converts character codes from 0x80 to 0x9f to Unicode equivalents
 * within a string
 * @param {string} string - the input string
 * @returns a converted string
 */
module.exports.binaryToUnicode = (string) => {
  return string.replace(/([\x80-\x9f])/g, (match) => binaryToUnicodeTable[match.charCodeAt(0)]);
};

/**
 * The main function for cleaning OLE-based text. It runs a few standard replacements on characters
 * that are reserved for special purposes, also removes fields, and finally strips out any weird 
 * characters that are likely not to be useful for anyone.
 * 
 * @param {string} string - an input string
 * @returns a cleaned up string
 */
module.exports.clean = (string) => {

  // Fields can be nested, which makes this awkward. We use a strict non-nesting model
  // and repeat until we find no substitutions. This is because a second match might
  // start before an earlier one, due to our replacements.

  string = string.replace(/([\x02\x05\x07\x08\x0a\x0b\x0c\x0d\x1f])/g, (match) => replaceTable[match.charCodeAt(0)]);

  let called = true;
  while (called) {
    called = false;
    string = string.replace(/(?:\x13[^\x13\x14\x15]*\x14?([^\x13\x14\x15]*)\x15)/g, (match, p1) => { called = true; return p1; });
  }

  return string
    .replace(/[\x00-\x07]/g, '');
};

const filterTable = [];
filterTable[0x2002] = " ";
filterTable[0x2003] = " ";
filterTable[0x2012] = "-";
filterTable[0x2013] = "-";
filterTable[0x2014] = "-";
filterTable[0x2018] = "'";
filterTable[0x2019] = "'";
filterTable[0x201c] = "\"";
filterTable[0x201d] = "\"";

/**
 * Filters a string, with a few common Unicode replacements, primarily for standard
 * punctuation like non-breaking spaces, hyphens, and left and right curly quotes.
 * @param {string} string - the input string
 * @returns a filtered string
 */
module.exports.filter = (string) => {
  return string
    .replace(/[\u2002\u2003\u2012\u2013\u2014\u2018\u2019\u201c\u201d]/g, (match) => filterTable[match.charCodeAt(0)]);
};

},{}],9:[function(require,module,exports){

const ALLOCATION_TABLE_SEC_ID_FREE               = -1;
const ALLOCATION_TABLE_SEC_ID_END_OF_CHAIN       = -2;   // eslint-disable-line no-unused-vars
const ALLOCATION_TABLE_SEC_ID_SAT                = -3;   // eslint-disable-line no-unused-vars
const ALLOCATION_TABLE_SEC_ID_MSAT               = -4;   // eslint-disable-line no-unused-vars

class AllocationTable {

  constructor(doc) {
    this._doc = doc;
  }

  load(secIds) {
    const doc = this._doc;
    const header = doc._header;
    this._table = new Array(secIds.length * (header.secSize / 4));
    return doc._readSectors(secIds)
      .then((buffer) => {
        for (let i = 0; i < buffer.length / 4; i++) {
          this._table[i] = buffer.readInt32LE(i * 4);
        }
      });
  }
  
  getSecIdChain(startSecId) {
    let secId = startSecId;
    const secIds = [];
    while (secId > ALLOCATION_TABLE_SEC_ID_FREE) {
      secIds.push(secId);
      const secIdPrior = secId;
      secId = this._table[secId];
      if (secId === secIdPrior) { // this will cause a deadlock and a out of memory error
        break;
      }
    }
  
    return secIds;
  }
  
}

module.exports = AllocationTable;

},{}],10:[function(require,module,exports){
(function (Buffer){(function (){
/**
 * @module ole-compound-doc
 */

// Copyright (c) 2012 Chris Geiersbach
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//
// This component as adapted from node-ole-doc, available at:
// https://github.com/atariman486/node-ole-doc.
//
// Modified extensively by Stuart Watt <stuart@morungos.com> to keep the 
// principal logic, but replacing callbacks and some weird stream usages
// with promises.

const Header = require('./ole-header');
const AllocationTable = require('./ole-allocation-table');
const DirectoryTree = require('./ole-directory-tree');
const Storage = require('./ole-storage');

/**
 * Implements the main interface used to read from an OLE compoound file.
 */
class OleCompoundDoc {

  constructor(reader) {    
    this._reader = reader;
    this._skipBytes = 0;
  }

  read() {
    return Promise.resolve()
      .then(() => this._readHeader())
      .then(() => this._readMSAT())
      .then(() => this._readSAT())
      .then(() => this._readSSAT())
      .then(() => this._readDirectoryTree())
      .then(() => {
        if (this._skipBytes != 0) {
          return this._readCustomHeader();
        }
      })
      .then(() => this);
  }

  _readCustomHeader() {
    const buffer = Buffer.alloc(this._skipBytes);
    return this._reader.read(buffer, 0, this._skipBytes, 0)
      .then((buffer) => {
        if (!this._customHeaderCallback(buffer))
          return;
      });
  }

  _readHeader() {
    const buffer = Buffer.alloc(512);
    return this._reader.read(buffer, 0, 512, 0 + this._skipBytes)
      .then((buffer) => {
        const header = this._header = new Header();
        if (!header.load(buffer)) {
          throw new Error("Not a valid compound document");
        }
      });
  }

  _readMSAT() {
    const header = this._header;

    this._MSAT = header.partialMSAT.slice(0);
    this._MSAT.length = header.SATSize;

    if(header.SATSize <= 109 || header.MSATSize == 0) {
      return Promise.resolve();
    }

    let currMSATIndex = 109;
    let i = 0;

    const readOneMSAT = (i, currMSATIndex, secId) => {
      if (i >= header.MSATSize) {
        return Promise.resolve();
      }

      return this._readSector(secId)
        .then((sectorBuffer) => {
          let s;
          for(s = 0; s < header.secSize - 4; s += 4) {
            if(currMSATIndex >= header.SATSize)
              break;
            else
              this._MSAT[currMSATIndex] = sectorBuffer.readInt32LE(s);

            currMSATIndex++;
          }

          secId = sectorBuffer.readInt32LE(header.secSize - 4);
          return readOneMSAT(i + 1, currMSATIndex, secId);
        });
    };

    return readOneMSAT(i, currMSATIndex, header.MSATSecId);
  }

  _readSector(secId) {
    return this._readSectors([ secId ]);
  }

  _readSectors(secIds) {
    const header = this._header;
    const buffer = Buffer.alloc(secIds.length * header.secSize);

    const readOneSector = (i) => {
      if (i >= secIds.length) {
        return Promise.resolve(buffer);
      }

      const bufferOffset = i * header.secSize;
      const fileOffset = this._getFileOffsetForSec(secIds[i]);

      return this._reader.read(buffer, bufferOffset, header.secSize, fileOffset)
        .then(() => readOneSector(i + 1));
    };

    return readOneSector(0);
  }

  _readShortSector(secId) {
    return this._readShortSectors([ secId ]);
  }

  _readShortSectors(secIds) {
    const header = this._header;
    const buffer = Buffer.alloc(secIds.length * header.shortSecSize);

    const readOneShortSector = (i) => {
      if (i >= secIds.length) {
        return Promise.resolve(buffer);
      }

      const bufferOffset = i * header.shortSecSize;
      const fileOffset = this._getFileOffsetForShortSec(secIds[i]);

      return this._reader.read(buffer, bufferOffset, header.shortSecSize, fileOffset)
        .then(() => readOneShortSector(i + 1));
    };

    return readOneShortSector(0);
  }

  _readSAT() {
    this._SAT = new AllocationTable(this);
    return this._SAT.load(this._MSAT);
  }

  _readSSAT() {
    const header = this._header;

    const secIds = this._SAT.getSecIdChain(header.SSATSecId);
    if (secIds.length != header.SSATSize) {
      return Promise.reject(new Error("Invalid Short Sector Allocation Table"));
    }

    this._SSAT = new AllocationTable(this);
    return this._SSAT.load(secIds);
  }

  _readDirectoryTree() {
    const header = this._header;

    this._directoryTree = new DirectoryTree(this);

    const secIds = this._SAT.getSecIdChain(header.dirSecId);
    return this._directoryTree.load(secIds)
      .then(() => {
        const rootEntry = this._directoryTree.root;
        this._rootStorage = new Storage(this, rootEntry);
        this._shortStreamSecIds = this._SAT.getSecIdChain(rootEntry.secId);
      });
  }

  _getFileOffsetForSec(secId) {
    const secSize = this._header.secSize;
    return this._skipBytes + (secId + 1) * secSize;  // Skip past the header sector
  }

  _getFileOffsetForShortSec(shortSecId) {
    const shortSecSize = this._header.shortSecSize;
    const shortStreamOffset = shortSecId * shortSecSize;

    const secSize = this._header.secSize;
    const secIdIndex = Math.floor(shortStreamOffset / secSize);
    const secOffset = shortStreamOffset % secSize;
    const secId = this._shortStreamSecIds[secIdIndex];

    return this._getFileOffsetForSec(secId) + secOffset;
  }

  storage(storageName) {
    return this._rootStorage.storage(storageName);
  }

  stream(streamName) {
    return this._rootStorage.stream(streamName);
  }

}

module.exports = OleCompoundDoc;

}).call(this)}).call(this,require("buffer").Buffer)
},{"./ole-allocation-table":9,"./ole-directory-tree":11,"./ole-header":12,"./ole-storage":14,"buffer":3}],11:[function(require,module,exports){
const DIRECTORY_TREE_ENTRY_TYPE_EMPTY =       0;   // eslint-disable-line no-unused-vars
const DIRECTORY_TREE_ENTRY_TYPE_STORAGE =     1;
const DIRECTORY_TREE_ENTRY_TYPE_STREAM =      2;
const DIRECTORY_TREE_ENTRY_TYPE_ROOT =        5;

const DIRECTORY_TREE_NODE_COLOR_RED =         0;   // eslint-disable-line no-unused-vars
const DIRECTORY_TREE_NODE_COLOR_BLACK =       1;   // eslint-disable-line no-unused-vars

const DIRECTORY_TREE_LEAF =                  -1;

class DirectoryTree {

  constructor(doc) {
    this._doc = doc;
  }

  load(secIds) {
    const doc = this._doc;

    return doc._readSectors(secIds)
      .then((buffer) => {
        const count = buffer.length / 128;
        this._entries = new Array(count);
        for(let i = 0; i < count; i++) {
          const offset = i * 128;
          const nameLength = Math.max(buffer.readInt16LE(64 + offset) - 2, 0);
  
          const entry = {};
          entry.name = buffer.toString('utf16le', 0 + offset, nameLength + offset);
          entry.type = buffer.readInt8(66 + offset);
          entry.nodeColor = buffer.readInt8(67 + offset);
          entry.left = buffer.readInt32LE(68 + offset);
          entry.right = buffer.readInt32LE(72 + offset);
          entry.storageDirId = buffer.readInt32LE(76 + offset);
          entry.secId = buffer.readInt32LE(116 + offset);
          entry.size = buffer.readInt32LE(120 + offset);
    
          this._entries[i] = entry;
        }
    
        this.root = this._entries.find((entry) => entry.type === DIRECTORY_TREE_ENTRY_TYPE_ROOT);
        this._buildHierarchy(this.root);
      });
  }

  _buildHierarchy(storageEntry) {
    const childIds = this._getChildIds(storageEntry);
  
    storageEntry.storages = {};
    storageEntry.streams  = {};

    for(const childId of childIds) {
      const childEntry = this._entries[childId];
      const name = childEntry.name;
      if (childEntry.type === DIRECTORY_TREE_ENTRY_TYPE_STORAGE) {
        storageEntry.storages[name] = childEntry;
      }
      if (childEntry.type === DIRECTORY_TREE_ENTRY_TYPE_STREAM) {
        storageEntry.streams[name] = childEntry;
      }
    }

    for(const name in storageEntry.storages) {
      this._buildHierarchy(storageEntry.storages[name]);
    }
  }
  
  _getChildIds(storageEntry) {
    const childIds = [];
  
    const visit = (visitEntry) => {
      if (visitEntry.left !== DIRECTORY_TREE_LEAF) {
        childIds.push(visitEntry.left);
        visit(this._entries[visitEntry.left]);
      }
      if (visitEntry.right !== DIRECTORY_TREE_LEAF) {
        childIds.push(visitEntry.right);
        visit(this._entries[visitEntry.right]);
      }
    };
  
    if (storageEntry.storageDirId > -1) {
      childIds.push(storageEntry.storageDirId);
      const rootChildEntry = this._entries[storageEntry.storageDirId];
      visit(rootChildEntry);
    }
  
    return childIds;
  }

}

module.exports = DirectoryTree;

},{}],12:[function(require,module,exports){
(function (Buffer){(function (){
const HEADER_DATA = Buffer.from('D0CF11E0A1B11AE1', 'hex');

class Header {

  constructor() {}

  load(buffer) {
    for(let i = 0; i < HEADER_DATA.length; i++) {
      if (HEADER_DATA[i] != buffer[i])
        return false;
    }
  
    this.secSize        = 1 << buffer.readInt16LE(30);  // Size of sectors
    this.shortSecSize   = 1 << buffer.readInt16LE(32);  // Size of short sectors
    this.SATSize        =      buffer.readInt32LE(44);  // Number of sectors used for the Sector Allocation Table
    this.dirSecId       =      buffer.readInt32LE(48);  // Starting Sec ID of the directory stream
    this.shortStreamMax =      buffer.readInt32LE(56);  // Maximum size of a short stream
    this.SSATSecId      =      buffer.readInt32LE(60);  // Starting Sec ID of the Short Sector Allocation Table
    this.SSATSize       =      buffer.readInt32LE(64);  // Number of sectors used for the Short Sector Allocation Table
    this.MSATSecId      =      buffer.readInt32LE(68);  // Starting Sec ID of the Master Sector Allocation Table
    this.MSATSize       =      buffer.readInt32LE(72);  // Number of sectors used for the Master Sector Allocation Table
  
    // The first 109 sectors of the MSAT
    this.partialMSAT = new Array(109);
    for(let i = 0; i < 109; i++)
      this.partialMSAT[i] = buffer.readInt32LE(76 + i * 4);
  
    return true;  
  }

}

module.exports = Header;

}).call(this)}).call(this,require("buffer").Buffer)
},{"buffer":3}],13:[function(require,module,exports){
const { EventEmitter } = require('events');

class StorageStream extends EventEmitter {

  constructor(doc, streamEntry) {
    super();
    this._doc = doc;
    this._streamEntry = streamEntry;
    this.initialize();
    setTimeout(() => this._pump(), 0);
  }

  initialize() {
    this._index = 0;
    this._done = true;

    if (!this._streamEntry) {
      return;
    }

    const doc  = this._doc;
    this._bytes = this._streamEntry.size;
  
    this._allocationTable = doc._SAT;
    this._shortStream = false;
    if (this._bytes < doc._header.shortStreamMax) {
      this._shortStream = true;
      this._allocationTable = doc._SSAT;
    }
  
    this._secIds = this._allocationTable.getSecIdChain(this._streamEntry.secId);
    this._done = false;
  }

  _readSector(sector) {
    if (this._shortStream) {
      return this._doc._readShortSector(sector);
    } else {
      return this._doc._readSector(sector);
    }
  }

  async _pump() {
    try {
      while (!this._done && this._index < this._secIds.length) {
        let buffer = await this._readSector(this._secIds[this._index]);
        if (this._bytes - buffer.length < 0) {
          buffer = buffer.slice(0, this._bytes);
        }

        this._bytes -= buffer.length;
        this._index ++;
        this.emit('data', buffer);
      }
      this._done = true;
      this.emit('end');
    } catch (error) {
      this.emit('error', error);
    }
  }
}

module.exports = StorageStream;

},{"events":4}],14:[function(require,module,exports){
const StorageStream = require('./ole-storage-stream');

class Storage {

  constructor(doc, dirEntry) {
    this._doc = doc;
    this._dirEntry = dirEntry;
  }

  storage(storageName) {
    return new Storage(this._doc, this._dirEntry.storages[storageName]);
  }
  
  stream(streamName) {
    return new StorageStream(this._doc, this._dirEntry.streams[streamName]);
  }
  
}

module.exports = Storage;
},{"./ole-storage-stream":13}],15:[function(require,module,exports){
(function (Buffer){(function (){
/**
 * @module word-ole-extractor
 * 
 * @description
 * Implements the main logic of extracting text from "classic" OLE-based Word files.
 * Depends on [OleCompoundDoc]{@link module:ole-compound-doc~OleCompoundDoc} 
 * for most of the underlying OLE logic. Note that
 * [OpenOfficeExtractor]{@link module:open-office-extractor~OpenOfficeExtractor} is 
 * used for newer, Open Office-style, files. 
 */

const OleCompoundDoc = require('./ole-compound-doc');
const Document = require('./document');
const { binaryToUnicode, clean } = require('./filters');

/**
 * Constant for the deletion character SPRM.
 */
const sprmCFRMarkDel = 0x00;

/**
 * Given a cp-style file offset, finds the containing piece index.
 * @param {*} offset the character offset
 * @returns the piece index
 * 
 * @todo 
 * Might be better using a binary search
 */
const getPieceIndexByCP = (pieces, position) => {
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    if (position <= piece.endCp) { 
      return i; 
    }
  }
};

/**
 * Given a file-style offset, finds the containing piece index.
 * @param {*} offset the character offset
 * @returns the piece index
 * 
 * @todo 
 * Might be better using a binary search
 */
const getPieceIndexByFilePos = (pieces, position) => {
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    if (position <= piece.endFilePos) { 
      return i; 
    }
  }
};

/**
 * Reads and extracts a character range from the pieces. This returns the 
 * plain text within the pieces in the given range.
 * @param {*} start the start offset
 * @param {*} end the end offset
 * @returns a character string
 */
function getTextRangeByCP(pieces, start, end) {
  const startPiece = getPieceIndexByCP(pieces, start);
  const endPiece = getPieceIndexByCP(pieces, end);
  const result = [];
  for (let i = startPiece, end1 = endPiece; i <= end1; i++) {
    const piece = pieces[i];
    const xstart = i === startPiece ? start - piece.startCp : 0;
    const xend = i === endPiece ? end - piece.startCp : piece.endCp;
    result.push(piece.text.substring(xstart, xend));
  }

  return result.join("");
}


/**
 * Given a piece, and a starting and ending cp-style file offset, 
 * and a replacement character, updates the piece text to replace
 * between start and end with the given character.
 * @param {*} piece the piece
 * @param {*} start the starting character offset
 * @param {*} end the endingcharacter offset
 * @param {*} character the replacement character
 */
function fillPieceRange(piece, start, end, character) {
  const pieceStart = piece.startCp;
  const pieceEnd = pieceStart + piece.length;
  const original = piece.text;
  if (start < pieceStart) start = pieceStart;
  if (end > pieceEnd) end = pieceEnd;
  const modified = 
    ((start == pieceStart) ? '' : original.slice(0, start - pieceStart)) +
    ''.padStart(end - start, character) +
    ((end == pieceEnd) ? '' : original.slice(end - pieceEnd));
  piece.text = modified;
}

/**
 * Given a piece, and a starting and ending filePos-style file offset, 
 * and a replacement character, updates the piece text to replace
 * between start and end with the given character. This is used when
 * applying character styles, which use filePos values rather than cp
 * values.
 *
 * @param {*} piece the piece
 * @param {*} start the starting character offset
 * @param {*} end the endingcharacter offset
 * @param {*} character the replacement character
 */
function fillPieceRangeByFilePos(piece, start, end, character) {
  const pieceStart = piece.startFilePos;
  const pieceEnd = pieceStart + piece.size;
  const original = piece.text;
  if (start < pieceStart) start = pieceStart;
  if (end > pieceEnd) end = pieceEnd;
  const modified = 
    ((start == pieceStart) ? '' : original.slice(0, (start - pieceStart) / piece.bpc)) +
    ''.padStart((end - start) / piece.bpc, character) +
    ((end == pieceEnd) ? '' : original.slice((end - pieceEnd) / piece.bpc));
  piece.text = modified;
}

/**
 * Replaces a selected range in the piece table, overwriting the selection with 
 * the given character. The length of segments in the piece table must never be
 * changed. 
 * @param {*} pieces 
 * @param {*} start 
 * @param {*} end 
 * @param {*} character 
 */
function replaceSelectedRange(pieces, start, end, character) {     // eslint-disable-line no-unused-vars
  const startPiece = getPieceIndexByCP(pieces, start);
  const endPiece = getPieceIndexByCP(pieces, end);
  for (let i = startPiece, end1 = endPiece; i <= end1; i++) {
    const piece = pieces[i];
    fillPieceRange(piece, start, end, character);
  }
}

/**
 * Replaces a selected range in the piece table, overwriting the selection with 
 * the given character. The length of segments in the piece table must never be
 * changed. The start and end values are found by file position.
 * @param {*} pieces 
 * @param {*} start 
 * @param {*} end 
 * @param {*} character 
 */
function replaceSelectedRangeByFilePos(pieces, start, end, character) {
  const startPiece = getPieceIndexByFilePos(pieces, start);
  const endPiece = getPieceIndexByFilePos(pieces, end);
  for (let i = startPiece, end1 = endPiece; i <= end1; i++) {
    const piece = pieces[i];
    fillPieceRangeByFilePos(piece, start, end, character);
  }
}

/**
 * Marks a range as deleted. It does this by overwriting it with null characters,
 * wich then get removed during the later cleaning process.
 * @param {*} pieces 
 * @param {*} start 
 * @param {*} end 
 */
function markDeletedRange(pieces, start, end) {
  replaceSelectedRangeByFilePos(pieces, start, end, '\x00');
}

/**
 * Called to iterate over a set of SPRMs in a buffer, starting at 
 * a gived offset. The handler is called with the arguments:
 * buffer, offset, sprm, ispmd, fspec, sgc, spra.
 * @param {*} buffer the buffer
 * @param {*} offset the starting offset
 * @param {*} handler the function to call for each SPRM
 */
const processSprms = (buffer, offset, handler) => {
  while (offset < buffer.length - 1) {
    const sprm = buffer.readUInt16LE(offset);
    const ispmd = sprm & 0x1f;
    const fspec = (sprm >> 9) & 0x01;
    const sgc = (sprm >> 10) & 0x07;
    const spra = (sprm >> 13) & 0x07;

    offset += 2;

    handler(buffer, offset, sprm, ispmd, fspec, sgc, spra);

    if (spra === 0) {
      offset += 1;
      continue;
    } else if (spra === 1) {
      offset += 1;
      continue;
    } else if (spra === 2) {
      offset += 2;
      continue;
    } else if (spra === 3) {
      offset += 4;
      continue;
    } else if (spra === 4 || spra === 5) {
      offset += 2;
      continue;
    } else if (spra === 6) {
      offset += buffer.readUInt8(offset) + 1;
      continue;
    } else if (spra === 7) {
      offset += 3;
      continue;
    } else {
      throw new Error("Unparsed sprm");
    }

  }
};

/**
 * @class
 * The main class implementing extraction from OLE-based Word files. 
 * This handles all the extraction and conversion logic. 
 */
class WordOleExtractor {

  constructor() { 
    this._pieces = [];
    this._bookmarks = {};
    this._boundaries = {};
    this._taggedHeaders = [];
  }

  /**
   * The main extraction method. This creates an OLE compound document
   * interface, then opens up a stream and extracts out the main
   * stream.
   * @param {*} reader 
   */
  extract(reader) {
    const document = new OleCompoundDoc(reader);
    return document.read()
      .then(() =>
        this.documentStream(document, 'WordDocument')
          .then((stream) => this.streamBuffer(stream))
          .then((buffer) => this.extractWordDocument(document, buffer))
      );
  }

  /**
   * Builds and returns a {@link Document} object corresponding to the text
   * in the original document. This involves reading and retrieving the text
   * ranges corresponding to the primary document parts. The text segments are
   * read from the extracted table of text pieces.
   * @returns a {@link Document} object
   */
  buildDocument() {
    const document = new Document();
    const pieces = this._pieces;

    let start = 0;

    document._body = clean(getTextRangeByCP(pieces, start, start + this._boundaries.ccpText));
    start += this._boundaries.ccpText;

    if (this._boundaries.ccpFtn) {
      document._footnotes = clean(getTextRangeByCP(pieces, start, start + this._boundaries.ccpFtn - 1));
      start += this._boundaries.ccpFtn;  
    }

    if (this._boundaries.ccpHdd) {
      // Replaced old single-block data with tagged selection. See #34
      // document._headers = clean(getTextRangeByCP(pieces, start, start + this._boundaries.ccpHdd - 1));
      document._headers = clean(this._taggedHeaders.filter((s) => s.type === 'headers').map((s) => s.text).join(""));
      document._footers = clean(this._taggedHeaders.filter((s) => s.type === 'footers').map((s) => s.text).join(""));

      start += this._boundaries.ccpHdd;  
    }

    if (this._boundaries.ccpAtn) {
      document._annotations = clean(getTextRangeByCP(pieces, start, start + this._boundaries.ccpAtn - 1));
      start += this._boundaries.ccpAtn;  
    }

    if (this._boundaries.ccpEdn) {
      document._endnotes = clean(getTextRangeByCP(pieces, start, start + this._boundaries.ccpEdn - 1));
      start += this._boundaries.ccpEdn;  
    }

    if (this._boundaries.ccpTxbx) {
      document._textboxes = clean(getTextRangeByCP(pieces, start, start + this._boundaries.ccpTxbx - 1));
      start += this._boundaries.ccpTxbx;  
    }

    if (this._boundaries.ccpHdrTxbx) {
      document._headerTextboxes = clean(getTextRangeByCP(pieces, start, start + this._boundaries.ccpHdrTxbx - 1));
      start += this._boundaries.ccpHdrTxbx;  
    }

    return document;
  }

  /**
   * Main logic top level function for unpacking a Word document 
   * @param {*} document the OLE document
   * @param {*} buffer a buffer 
   * @returns a Promise which resolves to a {@link Document}
   */
  extractWordDocument(document, buffer) {
    const magic = buffer.readUInt16LE(0);
    if (magic !== 0xa5ec) {
      return Promise.reject(new Error(`This does not seem to be a Word document: Invalid magic number: ${magic.toString(16)}`));
    }

    const flags = buffer.readUInt16LE(0xA);

    const streamName = (flags & 0x0200) !== 0 ? "1Table" : "0Table";

    return this.documentStream(document, streamName)
      .then((stream) => this.streamBuffer(stream))
      .then((streamBuffer) => {
        this._boundaries.fcMin = buffer.readUInt32LE(0x0018);
        this._boundaries.ccpText = buffer.readUInt32LE(0x004c);
        this._boundaries.ccpFtn = buffer.readUInt32LE(0x0050);
        this._boundaries.ccpHdd = buffer.readUInt32LE(0x0054);
        this._boundaries.ccpAtn = buffer.readUInt32LE(0x005c);
        this._boundaries.ccpEdn = buffer.readUInt32LE(0x0060);
        this._boundaries.ccpTxbx = buffer.readUInt32LE(0x0064);
        this._boundaries.ccpHdrTxbx = buffer.readUInt32LE(0x0068);

        this.writeBookmarks(buffer, streamBuffer);
        this.writePieces(buffer, streamBuffer);
        this.writeCharacterProperties(buffer, streamBuffer);
        this.writeParagraphProperties(buffer, streamBuffer);
        this.normalizeHeaders(buffer, streamBuffer);

        return this.buildDocument();
      });
  }

  /**
   * Returns a promise that resolves to the named stream.
   * @param {*} document 
   * @param {*} streamName 
   * @returns a promise that resolves to the named stream
   */
  documentStream(document, streamName) {
    return Promise.resolve(document.stream(streamName));
  }

  /**
   * Returns a promise that resolves to a Buffer containing the contents of 
   * the given stream. 
   * @param {*} stream 
   * @returns a promise that resolves to the sream contents
   */
  streamBuffer(stream) {
    return new Promise((resolve, reject) => {
      const chunks = [];
      stream.on('data', (chunk) => chunks.push(chunk));
      stream.on('error', (error) => reject(error));
      stream.on('end', () => resolve(Buffer.concat(chunks)));
      return stream;
    });
  }

  writeFields(buffer, tableBuffer, result) {    // eslint-disable-line no-unused-vars
    const fcPlcffldMom = buffer.readInt32LE(0x011a);
    const lcbPlcffldMom = buffer.readUInt32LE(0x011e);
    //console.log(fcPlcffldMom, lcbPlcffldMom, tableBuffer.length);

    if (lcbPlcffldMom == 0) {
      return;
    }

    const fieldCount = (lcbPlcffldMom - 4) / 6;
    //console.log("extracting", fieldCount, "fields");

    const dataOffset = (fieldCount + 1) * 4;

    const plcffldMom = tableBuffer.slice(fcPlcffldMom, fcPlcffldMom + lcbPlcffldMom);
    for(let i = 0; i < fieldCount; i++) {
      const cp = plcffldMom.readUInt32LE(i * 4);      // eslint-disable-line no-unused-vars
      const fld = plcffldMom.readUInt16LE(dataOffset + i * 2);
      const byte1 = fld & 0xff;
      const byte2 = fld >> 8;                         // eslint-disable-line no-unused-vars
      if ((byte1 & 0x1f) == 19) {
        //console.log("A", i, cp, byte1.toString(16), byte2.toString(16));
      } else {
        //console.log("B", i, cp, byte1.toString(16), byte2.toString(16));
      }
    }
  }

  /**
   * Extracts and stores the document bookmarks into a local field.
   * @param {*} buffer 
   * @param {*} tableBuffer 
   */
  writeBookmarks(buffer, tableBuffer) {
    const fcSttbfBkmk = buffer.readUInt32LE(0x0142);
    const lcbSttbfBkmk = buffer.readUInt32LE(0x0146);
    const fcPlcfBkf = buffer.readUInt32LE(0x014a);
    const lcbPlcfBkf = buffer.readUInt32LE(0x014e);
    const fcPlcfBkl = buffer.readUInt32LE(0x0152);
    const lcbPlcfBkl = buffer.readUInt32LE(0x0156);

    if (lcbSttbfBkmk === 0) { 
      return; 
    }

    const sttbfBkmk = tableBuffer.slice(fcSttbfBkmk, fcSttbfBkmk + lcbSttbfBkmk);
    const plcfBkf = tableBuffer.slice(fcPlcfBkf, fcPlcfBkf + lcbPlcfBkf);
    const plcfBkl = tableBuffer.slice(fcPlcfBkl, fcPlcfBkl + lcbPlcfBkl);

    const fcExtend = sttbfBkmk.readUInt16LE(0);
    const cData = sttbfBkmk.readUInt16LE(2);       // eslint-disable-line no-unused-vars
    const cbExtra = sttbfBkmk.readUInt16LE(4);     // eslint-disable-line no-unused-vars

    if (fcExtend !== 0xffff) {
      throw new Error("Internal error: unexpected single-byte bookmark data");
    }

    let offset = 6;
    const index = 0;

    while (offset < lcbSttbfBkmk) {
      let length = sttbfBkmk.readUInt16LE(offset);
      length = length * 2;
      const segment = sttbfBkmk.slice(offset + 2, offset + 2 + length);
      const cpStart = plcfBkf.readUInt32LE(index * 4);
      const cpEnd = plcfBkl.readUInt32LE(index * 4);
      this._bookmarks[segment] = {start: cpStart, end: cpEnd};
      offset = offset + length + 2;
    }
  }

  /**
   * Extracts and stores the document text pieces into a local field. This is
   * probably the most crucial part of text extraction, as it is where we
   * get text corresponding to character positions. These may be stored in a 
   * different order in the file compared to the order we want them. 
   * 
   * @param {*} buffer 
   * @param {*} tableBuffer 
   */
  writePieces(buffer, tableBuffer) {
    let flag;
    let pos = buffer.readUInt32LE(0x01a2);

    while (true) {                          // eslint-disable-line no-constant-condition
      flag = tableBuffer.readUInt8(pos);
      if (flag !== 1) { break; }

      pos = pos + 1;
      const skip = tableBuffer.readUInt16LE(pos);
      pos = pos + 2 + skip;
    }

    flag = tableBuffer.readUInt8(pos);
    pos = pos + 1;
    if (flag !== 2) {
      throw new Error("Internal error: ccorrupted Word file");
    }

    const pieceTableSize = tableBuffer.readUInt32LE(pos);
    pos = pos + 4;

    const pieces = (pieceTableSize - 4) / 12;

    let startCp = 0;
    let startStream = 0;

    for (let x = 0, end = pieces - 1; x <= end; x++) {
      const offset = pos + ((pieces + 1) * 4) + (x * 8) + 2;
      let startFilePos = tableBuffer.readUInt32LE(offset);
      let unicode = false;
      if ((startFilePos & 0x40000000) === 0) {
        unicode = true;
      } else {
        startFilePos = startFilePos & ~(0x40000000);
        startFilePos = Math.floor(startFilePos / 2);
      }
      const lStart = tableBuffer.readUInt32LE(pos + (x * 4));
      const lEnd = tableBuffer.readUInt32LE(pos + ((x + 1) * 4));
      const totLength = lEnd - lStart;

      const piece = {
        startCp,
        startStream,
        totLength,
        startFilePos,
        unicode,
        bpc: (unicode) ? 2 : 1
      };

      piece.size = piece.bpc * (lEnd - lStart); 

      const textBuffer = buffer.slice(startFilePos, startFilePos + piece.size);
      if (unicode) {
        piece.text = textBuffer.toString('ucs2');
      } else {
        piece.text = binaryToUnicode(textBuffer.toString('binary'));
      }
      
      piece.length = piece.text.length;

      piece.endCp = piece.startCp + piece.length;
      piece.endStream = piece.startStream + piece.size;
      piece.endFilePos = piece.startFilePos + piece.size;

      startCp = piece.endCp;
      startStream = piece.endStream;

      this._pieces.push(piece);
    }
  }

  /**
   * Processes the headers and footers. The main logic here is that we might have a mix 
   * of "real" and "pseudo" headers. For example, a footnote generates some footnote
   * separator footer elements, which, unless they contain something interesting, we 
   * can dispense with. In fact, we want to dispense with anything which is made up of
   * whitespace and control characters, in general. This means locating the segments of
   * text in the extracted pieces, and conditionally replacing them with nulls. 
   * 
   * @param {*} buffer 
   * @param {*} tableBuffer 
   */
  normalizeHeaders(buffer, tableBuffer) {
    const pieces = this._pieces;
    
    const fcPlcfhdd = buffer.readUInt32LE(0x00f2);
    const lcbPlcfhdd = buffer.readUInt32LE(0x00f6);
    if (lcbPlcfhdd < 8) {
      return;
    }

    const offset = this._boundaries.ccpText + this._boundaries.ccpFtn;
    const ccpHdd = this._boundaries.ccpHdd;

    const plcHdd = tableBuffer.slice(fcPlcfhdd, fcPlcfhdd + lcbPlcfhdd);
    const plcHddCount = (lcbPlcfhdd / 4);
    let start = offset + plcHdd.readUInt32LE(0);
    for(let i = 1; i < plcHddCount; i++) {
      let end = offset + plcHdd.readUInt32LE(i * 4);
      if (end > offset + ccpHdd) { 
        end = offset + ccpHdd; 
      }
      const string = getTextRangeByCP(pieces, start, end);
      const story = i - 1;
      if ([0, 1, 2].includes(story)) {
        this._taggedHeaders.push({type: 'footnoteSeparators', text: string});
      } else if ([3, 4, 5].includes(story)) {
        this._taggedHeaders.push({type: 'endSeparators', text: string});
      } else if ([0, 1, 4].includes(story % 6)) {
        this._taggedHeaders.push({type: 'headers', text: string});
      } else if ([2, 3, 5].includes(story % 6)) {
        this._taggedHeaders.push({type: 'footers', text: string});
      }

      if (! /[^\r\n\u0002-\u0008]/.test(string)) {
        replaceSelectedRange(pieces, start, end, "\x00");
      } else {
        replaceSelectedRange(pieces, end - 1, end, "\x00");
      }

      start = end;     // eslint-disable-line no-unused-vars
    }

    // The last character can always be dropped, but we handle that later anyways. 
  }

  writeParagraphProperties(buffer, tableBuffer) {
    const pieces = this._pieces;

    const fcPlcfbtePapx = buffer.readUInt32LE(0x0102);
    const lcbPlcfbtePapx = buffer.readUInt32LE(0x0106);

    const plcBtePapxCount = (lcbPlcfbtePapx - 4) / 8;
    const dataOffset = (plcBtePapxCount + 1) * 4;
    const plcBtePapx = tableBuffer.slice(fcPlcfbtePapx, fcPlcfbtePapx + lcbPlcfbtePapx);

    for(let i = 0; i < plcBtePapxCount; i++) {
      const cp = plcBtePapx.readUInt32LE(i * 4);    // eslint-disable-line no-unused-vars
      const papxFkpBlock = plcBtePapx.readUInt32LE(dataOffset + i * 4);
      //console.log("paragraph property", cp, papxFkpBlock);

      const papxFkpBlockBuffer = buffer.slice(papxFkpBlock * 512, (papxFkpBlock + 1) * 512);
      //console.log("papxFkpBlockBuffer", papxFkpBlockBuffer);

      const crun = papxFkpBlockBuffer.readUInt8(511);
      //console.log("crun", crun);

      for(let j = 0; j < crun; j++) {
        const rgfc = papxFkpBlockBuffer.readUInt32LE(j * 4);
        const rgfcNext = papxFkpBlockBuffer.readUInt32LE((j + 1) * 4);

        const cbLocation = (crun + 1) * 4 + j * 13;
        const cbIndex = papxFkpBlockBuffer.readUInt8(cbLocation) * 2;

        const cb = papxFkpBlockBuffer.readUInt8(cbIndex);
        let grpPrlAndIstd = null;
        if (cb !== 0) {
          grpPrlAndIstd = papxFkpBlockBuffer.slice(cbIndex + 1, cbIndex + 1 + (2 * cb) - 1);
        } else {
          const cb2 = papxFkpBlockBuffer.readUInt8(cbIndex + 1);
          grpPrlAndIstd = papxFkpBlockBuffer.slice(cbIndex + 2, cbIndex + 2 + (2 * cb2));
        }
        //console.log("para; ", j, "rgfc=", rgfc, "rgfcNext=", rgfcNext, "grpPrlAndIstd=", grpPrlAndIstd);

        const istd = grpPrlAndIstd.readUInt16LE(0);    // eslint-disable-line no-unused-vars
        processSprms(grpPrlAndIstd, 2, (buffer, offset, sprm, ispmd, fspec, sgc, spra) => {    // eslint-disable-line no-unused-vars
          //console.log("sprm x", offset, sprm.toString(16), ispmd, fspec, sgc, spra);
          if (sprm === 0x2417) {
            replaceSelectedRangeByFilePos(pieces, rgfc, rgfcNext, '\n');
          }
        });
      }
      
    }
  }

  writeCharacterProperties(buffer, tableBuffer) {
    const pieces = this._pieces;

    const fcPlcfbteChpx = buffer.readUInt32LE(0x00fa);
    const lcbPlcfbteChpx = buffer.readUInt32LE(0x00fe);

    const plcBteChpxCount = (lcbPlcfbteChpx - 4) / 8;
    //console.log("character format runs", plcBteChpxCount, fcPlcfbteChpx, lcbPlcfbteChpx);

    const dataOffset = (plcBteChpxCount + 1) * 4;
    const plcBteChpx = tableBuffer.slice(fcPlcfbteChpx, fcPlcfbteChpx + lcbPlcfbteChpx);

    //const cpLast = plcBteChpx.readUInt32LE(plcBteChpxCount * 4);
    //console.log("last cp", cpLast);

    let lastDeletionEnd = null;

    for(let i = 0; i < plcBteChpxCount; i++) {
      const cp = plcBteChpx.readUInt32LE(i * 4);    // eslint-disable-line no-unused-vars
      const chpxFkpBlock = plcBteChpx.readUInt32LE(dataOffset + i * 4);
      //console.log("character property", cp, chpxFkpBlock);

      const chpxFkpBlockBuffer = buffer.slice(chpxFkpBlock * 512, (chpxFkpBlock + 1) * 512);
      //console.log("chpxFkpBlockBuffer", chpxFkpBlockBuffer);

      const crun = chpxFkpBlockBuffer.readUInt8(511);
      //console.log("crun", crun);

      for(let j = 0; j < crun; j++) {
        const rgfc = chpxFkpBlockBuffer.readUInt32LE(j * 4);
        const rgfcNext = chpxFkpBlockBuffer.readUInt32LE((j + 1) * 4);
        const rgb = chpxFkpBlockBuffer.readUInt8((crun + 1) * 4 + j);
        if (rgb == 0) {
          //console.log("skipping run; ", j, "rgfc=", rgfc, "rgb=", rgb);
          continue;
        }
        const chpxOffset = rgb * 2;
        const cb = chpxFkpBlockBuffer.readUInt8(chpxOffset);
        const grpprl = chpxFkpBlockBuffer.slice(chpxOffset + 1, chpxOffset + 1 + cb);
        //console.log("found run; ", j, "rgfc=", rgfc, "rgb=", rgb, "cb=", cb, "grpprl=", grpprl);

        processSprms(grpprl, 0, (buffer, offset, sprm, ispmd) => {
          if (ispmd === sprmCFRMarkDel) {
            if ((buffer[offset] & 1) != 1) {
              return;
            }

            // console.log("text deleted", rgfc, rgfcNext);
            if (lastDeletionEnd === rgfc) {
              markDeletedRange(pieces, lastDeletionEnd, rgfcNext);
            } else {
              markDeletedRange(pieces, rgfc, rgfcNext);
            }
            lastDeletionEnd = rgfcNext;

            // if (ld >= 0 && this._deletions[ld].end === rgfc) {
            //   this._deletions[ld].end = rgfcNext;
            // } else {
            //   this._deletions.push({start: rgfc, end: rgfcNext});
            // }
          }
        });
      }
    }
  }

}

module.exports = WordOleExtractor;

}).call(this)}).call(this,require("buffer").Buffer)
},{"./document":7,"./filters":8,"./ole-compound-doc":10,"buffer":3}]},{},[1])(1)
});
