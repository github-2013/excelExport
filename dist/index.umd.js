
(function(l, r) { if (!l || l.getElementById('livereloadscript')) return; r = l.createElement('script'); r.async = 1; r.src = '//' + (self.location.host || 'localhost').split(':')[0] + ':35729/livereload.js?snipver=1'; r.id = 'livereloadscript'; l.getElementsByTagName('head')[0].appendChild(r) })(self.document);
(function (global, factory) {
	typeof exports === 'object' && typeof module !== 'undefined' ? module.exports = factory() :
	typeof define === 'function' && define.amd ? define(factory) :
	(global = typeof globalThis !== 'undefined' ? globalThis : global || self, global.json2excel = factory());
})(this, (function () { 'use strict';

	/*! xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
	/* vim: set ts=2: */
	/*exported XLSX */
	/*global process:false, Buffer:false, ArrayBuffer:false, DataView:false, Deno:false */

	var _getchar = function _gc1(x/*:number*/)/*:string*/ { return String.fromCharCode(x); };
	var Base64_map = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
	function Base64_encode(input) {
	  var o = "";
	  var c1 = 0, c2 = 0, c3 = 0, e1 = 0, e2 = 0, e3 = 0, e4 = 0;
	  for (var i = 0; i < input.length; ) {
	    c1 = input.charCodeAt(i++);
	    e1 = c1 >> 2;
	    c2 = input.charCodeAt(i++);
	    e2 = (c1 & 3) << 4 | c2 >> 4;
	    c3 = input.charCodeAt(i++);
	    e3 = (c2 & 15) << 2 | c3 >> 6;
	    e4 = c3 & 63;
	    if (isNaN(c2)) {
	      e3 = e4 = 64;
	    } else if (isNaN(c3)) {
	      e4 = 64;
	    }
	    o += Base64_map.charAt(e1) + Base64_map.charAt(e2) + Base64_map.charAt(e3) + Base64_map.charAt(e4);
	  }
	  return o;
	}
	function Base64_decode(input) {
	  var o = "";
	  var c1 = 0, c2 = 0, c3 = 0, e1 = 0, e2 = 0, e3 = 0, e4 = 0;
	  input = input.replace(/[^\w\+\/\=]/g, "");
	  for (var i = 0; i < input.length; ) {
	    e1 = Base64_map.indexOf(input.charAt(i++));
	    e2 = Base64_map.indexOf(input.charAt(i++));
	    c1 = e1 << 2 | e2 >> 4;
	    o += String.fromCharCode(c1);
	    e3 = Base64_map.indexOf(input.charAt(i++));
	    c2 = (e2 & 15) << 4 | e3 >> 2;
	    if (e3 !== 64) {
	      o += String.fromCharCode(c2);
	    }
	    e4 = Base64_map.indexOf(input.charAt(i++));
	    c3 = (e3 & 3) << 6 | e4;
	    if (e4 !== 64) {
	      o += String.fromCharCode(c3);
	    }
	  }
	  return o;
	}
	var has_buf = /*#__PURE__*/(function() { return typeof Buffer !== 'undefined' && typeof process !== 'undefined' && typeof process.versions !== 'undefined' && !!process.versions.node; })();

	var Buffer_from = /*#__PURE__*/(function() {
		if(typeof Buffer !== 'undefined') {
			var nbfs = !Buffer.from;
			if(!nbfs) try { Buffer.from("foo", "utf8"); } catch(e) { nbfs = true; }
			return nbfs ? function(buf, enc) { return (enc) ? new Buffer(buf, enc) : new Buffer(buf); } : Buffer.from.bind(Buffer);
		}
		return function() {};
	})();


	function new_raw_buf(len/*:number*/) {
		/* jshint -W056 */
		if(has_buf) return Buffer.alloc ? Buffer.alloc(len) : new Buffer(len);
		return typeof Uint8Array != "undefined" ? new Uint8Array(len) : new Array(len);
		/* jshint +W056 */
	}

	function new_unsafe_buf(len/*:number*/) {
		/* jshint -W056 */
		if(has_buf) return Buffer.allocUnsafe ? Buffer.allocUnsafe(len) : new Buffer(len);
		return typeof Uint8Array != "undefined" ? new Uint8Array(len) : new Array(len);
		/* jshint +W056 */
	}

	var s2a = function s2a(s/*:string*/)/*:any*/ {
		if(has_buf) return Buffer_from(s, "binary");
		return s.split("").map(function(x/*:string*/)/*:number*/{ return x.charCodeAt(0) & 0xff; });
	};

	function s2ab(s/*:string*/)/*:any*/ {
		if(typeof ArrayBuffer === 'undefined') return s2a(s);
		var buf = new ArrayBuffer(s.length), view = new Uint8Array(buf);
		for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
		return buf;
	}

	function a2s(data/*:any*/)/*:string*/ {
		if(Array.isArray(data)) return data.map(function(c) { return String.fromCharCode(c); }).join("");
		var o/*:Array<string>*/ = []; for(var i = 0; i < data.length; ++i) o[i] = String.fromCharCode(data[i]); return o.join("");
	}

	function a2u(data/*:Array<number>*/)/*:Uint8Array*/ {
		if(typeof Uint8Array === 'undefined') throw new Error("Unsupported");
		return new Uint8Array(data);
	}

	var bconcat = has_buf ? function(bufs) { return Buffer.concat(bufs.map(function(buf) { return Buffer.isBuffer(buf) ? buf : Buffer_from(buf); })); } : function(bufs) {
		if(typeof Uint8Array !== "undefined") {
			var i = 0, maxlen = 0;
			for(i = 0; i < bufs.length; ++i) maxlen += bufs[i].length;
			var o = new Uint8Array(maxlen);
			var len = 0;
			for(i = 0, maxlen = 0; i < bufs.length; maxlen += len, ++i) {
				len = bufs[i].length;
				if(bufs[i] instanceof Uint8Array) o.set(bufs[i], maxlen);
				else if(typeof bufs[i] == "string") { throw "wtf"; }
				else o.set(new Uint8Array(bufs[i]), maxlen);
			}
			return o;
		}
		return [].concat.apply([], bufs.map(function(buf) { return Array.isArray(buf) ? buf : [].slice.call(buf); }));
	};

	function utf8decode(content/*:string*/) {
		var out = [], widx = 0, L = content.length + 250;
		var o = new_raw_buf(content.length + 255);
		for(var ridx = 0; ridx < content.length; ++ridx) {
			var c = content.charCodeAt(ridx);
			if(c < 0x80) o[widx++] = c;
			else if(c < 0x800) {
				o[widx++] = (192|((c>>6)&31));
				o[widx++] = (128|(c&63));
			} else if(c >= 0xD800 && c < 0xE000) {
				c = (c&1023)+64;
				var d = content.charCodeAt(++ridx)&1023;
				o[widx++] = (240|((c>>8)&7));
				o[widx++] = (128|((c>>2)&63));
				o[widx++] = (128|((d>>6)&15)|((c&3)<<4));
				o[widx++] = (128|(d&63));
			} else {
				o[widx++] = (224|((c>>12)&15));
				o[widx++] = (128|((c>>6)&63));
				o[widx++] = (128|(c&63));
			}
			if(widx > L) {
				out.push(o.slice(0, widx));
				widx = 0;
				o = new_raw_buf(65535);
				L = 65530;
			}
		}
		out.push(o.slice(0, widx));
		return bconcat(out);
	}

	var chr0 = /\u0000/g, chr1 = /[\u0001-\u0006]/g;
	/*::
	declare type Block = any;
	declare type BufArray = {
		newblk(sz:number):Block;
		next(sz:number):Block;
		end():any;
		push(buf:Block):void;
	};

	type RecordHopperCB = {(d:any, Rn:string, RT:number):?boolean;};

	type EvertType = {[string]:string};
	type EvertNumType = {[string]:number};
	type EvertArrType = {[string]:Array<string>};

	type StringConv = {(string):string};

	*/
	/* ssf.js (C) 2013-present SheetJS -- http://sheetjs.com */
	/*jshint -W041 */
	function _strrev(x/*:string*/)/*:string*/ { var o = "", i = x.length-1; while(i>=0) o += x.charAt(i--); return o; }
	function pad0(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
	function pad_(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v;return t.length>=d?t:fill(' ',d-t.length)+t;}
	function rpad_(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v; return t.length>=d?t:t+fill(' ',d-t.length);}
	function pad0r1(v/*:any*/,d/*:number*/)/*:string*/{var t=""+Math.round(v); return t.length>=d?t:fill('0',d-t.length)+t;}
	function pad0r2(v/*:any*/,d/*:number*/)/*:string*/{var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
	var p2_32 = /*#__PURE__*/Math.pow(2,32);
	function pad0r(v/*:any*/,d/*:number*/)/*:string*/{if(v>p2_32||v<-p2_32) return pad0r1(v,d); var i = Math.round(v); return pad0r2(i,d); }
	/* yes, in 2022 this is still faster than string compare */
	function SSF_isgeneral(s/*:string*/, i/*:?number*/)/*:boolean*/ { i = i || 0; return s.length >= 7 + i && (s.charCodeAt(i)|32) === 103 && (s.charCodeAt(i+1)|32) === 101 && (s.charCodeAt(i+2)|32) === 110 && (s.charCodeAt(i+3)|32) === 101 && (s.charCodeAt(i+4)|32) === 114 && (s.charCodeAt(i+5)|32) === 97 && (s.charCodeAt(i+6)|32) === 108; }
	var days/*:Array<Array<string> >*/ = [
		['Sun', 'Sunday'],
		['Mon', 'Monday'],
		['Tue', 'Tuesday'],
		['Wed', 'Wednesday'],
		['Thu', 'Thursday'],
		['Fri', 'Friday'],
		['Sat', 'Saturday']
	];
	var months/*:Array<Array<string> >*/ = [
		['J', 'Jan', 'January'],
		['F', 'Feb', 'February'],
		['M', 'Mar', 'March'],
		['A', 'Apr', 'April'],
		['M', 'May', 'May'],
		['J', 'Jun', 'June'],
		['J', 'Jul', 'July'],
		['A', 'Aug', 'August'],
		['S', 'Sep', 'September'],
		['O', 'Oct', 'October'],
		['N', 'Nov', 'November'],
		['D', 'Dec', 'December']
	];
	function SSF_init_table(t/*:any*/) {
		if(!t) t = {};
		t[0]=  'General';
		t[1]=  '0';
		t[2]=  '0.00';
		t[3]=  '#,##0';
		t[4]=  '#,##0.00';
		t[9]=  '0%';
		t[10]= '0.00%';
		t[11]= '0.00E+00';
		t[12]= '# ?/?';
		t[13]= '# ??/??';
		t[14]= 'm/d/yy';
		t[15]= 'd-mmm-yy';
		t[16]= 'd-mmm';
		t[17]= 'mmm-yy';
		t[18]= 'h:mm AM/PM';
		t[19]= 'h:mm:ss AM/PM';
		t[20]= 'h:mm';
		t[21]= 'h:mm:ss';
		t[22]= 'm/d/yy h:mm';
		t[37]= '#,##0 ;(#,##0)';
		t[38]= '#,##0 ;[Red](#,##0)';
		t[39]= '#,##0.00;(#,##0.00)';
		t[40]= '#,##0.00;[Red](#,##0.00)';
		t[45]= 'mm:ss';
		t[46]= '[h]:mm:ss';
		t[47]= 'mmss.0';
		t[48]= '##0.0E+0';
		t[49]= '@';
		t[56]= '"上午/下午 "hh"時"mm"分"ss"秒 "';
		return t;
	}
	/* repeated to satiate webpack */
	var table_fmt = {
		0:  'General',
		1:  '0',
		2:  '0.00',
		3:  '#,##0',
		4:  '#,##0.00',
		9:  '0%',
		10: '0.00%',
		11: '0.00E+00',
		12: '# ?/?',
		13: '# ??/??',
		14: 'm/d/yy',
		15: 'd-mmm-yy',
		16: 'd-mmm',
		17: 'mmm-yy',
		18: 'h:mm AM/PM',
		19: 'h:mm:ss AM/PM',
		20: 'h:mm',
		21: 'h:mm:ss',
		22: 'm/d/yy h:mm',
		37: '#,##0 ;(#,##0)',
		38: '#,##0 ;[Red](#,##0)',
		39: '#,##0.00;(#,##0.00)',
		40: '#,##0.00;[Red](#,##0.00)',
		45: 'mm:ss',
		46: '[h]:mm:ss',
		47: 'mmss.0',
		48: '##0.0E+0',
		49: '@',
		56: '"上午/下午 "hh"時"mm"分"ss"秒 "'
	};

	/* Defaults determined by systematically testing in Excel 2019 */

	/* These formats appear to default to other formats in the table */
	var SSF_default_map = {
		5:  37, 6:  38, 7:  39, 8:  40,         //  5 -> 37 ...  8 -> 40

		23:  0, 24:  0, 25:  0, 26:  0,         // 23 ->  0 ... 26 ->  0

		27: 14, 28: 14, 29: 14, 30: 14, 31: 14, // 27 -> 14 ... 31 -> 14

		50: 14, 51: 14, 52: 14, 53: 14, 54: 14, // 50 -> 14 ... 58 -> 14
		55: 14, 56: 14, 57: 14, 58: 14,
		59:  1, 60:  2, 61:  3, 62:  4,         // 59 ->  1 ... 62 ->  4

		67:  9, 68: 10,                         // 67 ->  9 ... 68 -> 10
		69: 12, 70: 13, 71: 14,                 // 69 -> 12 ... 71 -> 14
		72: 14, 73: 15, 74: 16, 75: 17,         // 72 -> 14 ... 75 -> 17
		76: 20, 77: 21, 78: 22,                 // 76 -> 20 ... 78 -> 22
		79: 45, 80: 46, 81: 47,                 // 79 -> 45 ... 81 -> 47
		82: 0                                   // 82 ->  0 ... 65536 -> 0 (omitted)
	};


	/* These formats technically refer to Accounting formats with no equivalent */
	var SSF_default_str = {
		//  5 -- Currency,   0 decimal, black negative
		5:  '"$"#,##0_);\\("$"#,##0\\)',
		63: '"$"#,##0_);\\("$"#,##0\\)',

		//  6 -- Currency,   0 decimal, red   negative
		6:  '"$"#,##0_);[Red]\\("$"#,##0\\)',
		64: '"$"#,##0_);[Red]\\("$"#,##0\\)',

		//  7 -- Currency,   2 decimal, black negative
		7:  '"$"#,##0.00_);\\("$"#,##0.00\\)',
		65: '"$"#,##0.00_);\\("$"#,##0.00\\)',

		//  8 -- Currency,   2 decimal, red   negative
		8:  '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
		66: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',

		// 41 -- Accounting, 0 decimal, No Symbol
		41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',

		// 42 -- Accounting, 0 decimal, $  Symbol
		42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',

		// 43 -- Accounting, 2 decimal, No Symbol
		43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',

		// 44 -- Accounting, 2 decimal, $  Symbol
		44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
	};

	function SSF_frac(x/*:number*/, D/*:number*/, mixed/*:?boolean*/)/*:Array<number>*/ {
		var sgn = x < 0 ? -1 : 1;
		var B = x * sgn;
		var P_2 = 0, P_1 = 1, P = 0;
		var Q_2 = 1, Q_1 = 0, Q = 0;
		var A = Math.floor(B);
		while(Q_1 < D) {
			A = Math.floor(B);
			P = A * P_1 + P_2;
			Q = A * Q_1 + Q_2;
			if((B - A) < 0.00000005) break;
			B = 1 / (B - A);
			P_2 = P_1; P_1 = P;
			Q_2 = Q_1; Q_1 = Q;
		}
		if(Q > D) { if(Q_1 > D) { Q = Q_2; P = P_2; } else { Q = Q_1; P = P_1; } }
		if(!mixed) return [0, sgn * P, Q];
		var q = Math.floor(sgn * P/Q);
		return [q, sgn*P - q*Q, Q];
	}
	function SSF_parse_date_code(v/*:number*/,opts/*:?any*/,b2/*:?boolean*/) {
		if(v > 2958465 || v < 0) return null;
		var date = (v|0), time = Math.floor(86400 * (v - date)), dow=0;
		var dout=[];
		var out={D:date, T:time, u:86400*(v-date)-time,y:0,m:0,d:0,H:0,M:0,S:0,q:0};
		if(Math.abs(out.u) < 1e-6) out.u = 0;
		if(opts && opts.date1904) date += 1462;
		if(out.u > 0.9999) {
			out.u = 0;
			if(++time == 86400) { out.T = time = 0; ++date; ++out.D; }
		}
		if(date === 60) {dout = b2 ? [1317,10,29] : [1900,2,29]; dow=3;}
		else if(date === 0) {dout = b2 ? [1317,8,29] : [1900,1,0]; dow=6;}
		else {
			if(date > 60) --date;
			/* 1 = Jan 1 1900 in Gregorian */
			var d = new Date(1900, 0, 1);
			d.setDate(d.getDate() + date - 1);
			dout = [d.getFullYear(), d.getMonth()+1,d.getDate()];
			dow = d.getDay();
			if(date < 60) dow = (dow + 6) % 7;
			if(b2) dow = SSF_fix_hijri(d, dout);
		}
		out.y = dout[0]; out.m = dout[1]; out.d = dout[2];
		out.S = time % 60; time = Math.floor(time / 60);
		out.M = time % 60; time = Math.floor(time / 60);
		out.H = time;
		out.q = dow;
		return out;
	}
	var SSFbasedate = /*#__PURE__*/new Date(1899, 11, 31, 0, 0, 0);
	var SSFdnthresh = /*#__PURE__*/SSFbasedate.getTime();
	var SSFbase1904 = /*#__PURE__*/new Date(1900, 2, 1, 0, 0, 0);
	function datenum_local(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
		var epoch = /*#__PURE__*/v.getTime();
		if(date1904) epoch -= 1461*24*60*60*1000;
		else if(v >= SSFbase1904) epoch += 24*60*60*1000;
		return (epoch - (SSFdnthresh + (/*#__PURE__*/v.getTimezoneOffset() - /*#__PURE__*/SSFbasedate.getTimezoneOffset()) * 60000)) / (24 * 60 * 60 * 1000);
	}
	/* ECMA-376 18.8.30 numFmt*/
	/* Note: `toPrecision` uses standard form when prec > E and E >= -6 */
	/* exponent >= -9 and <= 9 */
	function SSF_strip_decimal(o/*:string*/)/*:string*/ {
		return (o.indexOf(".") == -1) ? o : o.replace(/(?:\.0*|(\.\d*[1-9])0+)$/, "$1");
	}

	/* General Exponential always shows 2 digits exp and trims the mantissa */
	function SSF_normalize_exp(o/*:string*/)/*:string*/ {
		if(o.indexOf("E") == -1) return o;
		return o.replace(/(?:\.0*|(\.\d*[1-9])0+)[Ee]/,"$1E").replace(/(E[+-])(\d)$/,"$10$2");
	}

	/* exponent >= -9 and <= 9 */
	function SSF_small_exp(v/*:number*/)/*:string*/ {
		var w = (v<0?12:11);
		var o = SSF_strip_decimal(v.toFixed(12)); if(o.length <= w) return o;
		o = v.toPrecision(10); if(o.length <= w) return o;
		return v.toExponential(5);
	}

	/* exponent >= 11 or <= -10 likely exponential */
	function SSF_large_exp(v/*:number*/)/*:string*/ {
		var o = SSF_strip_decimal(v.toFixed(11));
		return (o.length > (v<0?12:11) || o === "0" || o === "-0") ? v.toPrecision(6) : o;
	}

	function SSF_general_num(v/*:number*/)/*:string*/ {
		var V = Math.floor(Math.log(Math.abs(v))*Math.LOG10E), o;

		if(V >= -4 && V <= -1) o = v.toPrecision(10+V);
		else if(Math.abs(V) <= 9) o = SSF_small_exp(v);
		else if(V === 10) o = v.toFixed(10).substr(0,12);
		else o = SSF_large_exp(v);

		return SSF_strip_decimal(SSF_normalize_exp(o.toUpperCase()));
	}


	/*
		"General" rules:
		- text is passed through ("@")
		- booleans are rendered as TRUE/FALSE
		- "up to 11 characters" displayed for numbers
		- Default date format (code 14) used for Dates

		The longest 32-bit integer text is "-2147483648", exactly 11 chars
		TODO: technically the display depends on the width of the cell
	*/
	function SSF_general(v/*:any*/, opts/*:any*/) {
		switch(typeof v) {
			case 'string': return v;
			case 'boolean': return v ? "TRUE" : "FALSE";
			case 'number': return (v|0) === v ? v.toString(10) : SSF_general_num(v);
			case 'undefined': return "";
			case 'object':
				if(v == null) return "";
				if(v instanceof Date) return SSF_format(14, datenum_local(v, opts && opts.date1904), opts);
		}
		throw new Error("unsupported value in General format: " + v);
	}

	function SSF_fix_hijri(date/*:Date*/, o/*:[number, number, number]*/) {
	  /* TODO: properly adjust y/m/d and  */
	  o[0] -= 581;
	  var dow = date.getDay();
	  if(date < 60) dow = (dow + 6) % 7;
	  return dow;
	}
	//var THAI_DIGITS = "\u0E50\u0E51\u0E52\u0E53\u0E54\u0E55\u0E56\u0E57\u0E58\u0E59".split("");
	function SSF_write_date(type/*:number*/, fmt/*:string*/, val, ss0/*:?number*/)/*:string*/ {
		var o="", ss=0, tt=0, y = val.y, out, outl = 0;
		switch(type) {
			case 98: /* 'b' buddhist year */
				y = val.y + 543;
				/* falls through */
			case 121: /* 'y' year */
			switch(fmt.length) {
				case 1: case 2: out = y % 100; outl = 2; break;
				default: out = y % 10000; outl = 4; break;
			} break;
			case 109: /* 'm' month */
			switch(fmt.length) {
				case 1: case 2: out = val.m; outl = fmt.length; break;
				case 3: return months[val.m-1][1];
				case 5: return months[val.m-1][0];
				default: return months[val.m-1][2];
			} break;
			case 100: /* 'd' day */
			switch(fmt.length) {
				case 1: case 2: out = val.d; outl = fmt.length; break;
				case 3: return days[val.q][0];
				default: return days[val.q][1];
			} break;
			case 104: /* 'h' 12-hour */
			switch(fmt.length) {
				case 1: case 2: out = 1+(val.H+11)%12; outl = fmt.length; break;
				default: throw 'bad hour format: ' + fmt;
			} break;
			case 72: /* 'H' 24-hour */
			switch(fmt.length) {
				case 1: case 2: out = val.H; outl = fmt.length; break;
				default: throw 'bad hour format: ' + fmt;
			} break;
			case 77: /* 'M' minutes */
			switch(fmt.length) {
				case 1: case 2: out = val.M; outl = fmt.length; break;
				default: throw 'bad minute format: ' + fmt;
			} break;
			case 115: /* 's' seconds */
				if(fmt != 's' && fmt != 'ss' && fmt != '.0' && fmt != '.00' && fmt != '.000') throw 'bad second format: ' + fmt;
				if(val.u === 0 && (fmt == "s" || fmt == "ss")) return pad0(val.S, fmt.length);
				/*::if(!ss0) ss0 = 0; */
				if(ss0 >= 2) tt = ss0 === 3 ? 1000 : 100;
				else tt = ss0 === 1 ? 10 : 1;
				ss = Math.round((tt)*(val.S + val.u));
				if(ss >= 60*tt) ss = 0;
				if(fmt === 's') return ss === 0 ? "0" : ""+ss/tt;
				o = pad0(ss,2 + ss0);
				if(fmt === 'ss') return o.substr(0,2);
				return "." + o.substr(2,fmt.length-1);
			case 90: /* 'Z' absolute time */
			switch(fmt) {
				case '[h]': case '[hh]': out = val.D*24+val.H; break;
				case '[m]': case '[mm]': out = (val.D*24+val.H)*60+val.M; break;
				case '[s]': case '[ss]': out = ((val.D*24+val.H)*60+val.M)*60+Math.round(val.S+val.u); break;
				default: throw 'bad abstime format: ' + fmt;
			} outl = fmt.length === 3 ? 1 : 2; break;
			case 101: /* 'e' era */
				out = y; outl = 1; break;
		}
		var outstr = outl > 0 ? pad0(out, outl) : "";
		return outstr;
	}


	/*jshint -W086 */
	/*jshint +W086 */
	function commaify(s/*:string*/)/*:string*/ {
		var w = 3;
		if(s.length <= w) return s;
		var j = (s.length % w), o = s.substr(0,j);
		for(; j!=s.length; j+=w) o+=(o.length > 0 ? "," : "") + s.substr(j,w);
		return o;
	}
	var pct1 = /%/g;
	function write_num_pct(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/{
		var sfmt = fmt.replace(pct1,""), mul = fmt.length - sfmt.length;
		return write_num(type, sfmt, val * Math.pow(10,2*mul)) + fill("%",mul);
	}

	function write_num_cm(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/{
		var idx = fmt.length - 1;
		while(fmt.charCodeAt(idx-1) === 44) --idx;
		return write_num(type, fmt.substr(0,idx), val / Math.pow(10,3*(fmt.length-idx)));
	}

	function write_num_exp(fmt/*:string*/, val/*:number*/)/*:string*/{
		var o/*:string*/;
		var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
		if(fmt.match(/^#+0.0E\+0$/)) {
			if(val == 0) return "0.0E+0";
			else if(val < 0) return "-" + write_num_exp(fmt, -val);
			var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
			var ee = Math.floor(Math.log(val)*Math.LOG10E)%period;
			if(ee < 0) ee += period;
			o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
			if(o.indexOf("e") === -1) {
				var fakee = Math.floor(Math.log(val)*Math.LOG10E);
				if(o.indexOf(".") === -1) o = o.charAt(0) + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
				else o += "E+" + (fakee - ee);
				while(o.substr(0,2) === "0.") {
					o = o.charAt(0) + o.substr(2,period) + "." + o.substr(2+period);
					o = o.replace(/^0+([1-9])/,"$1").replace(/^0+\./,"0.");
				}
				o = o.replace(/\+-/,"-");
			}
			o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
		} else o = val.toExponential(idx);
		if(fmt.match(/E\+00$/) && o.match(/e[+-]\d$/)) o = o.substr(0,o.length-1) + "0" + o.charAt(o.length-1);
		if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
		return o.replace("e","E");
	}
	var frac1 = /# (\?+)( ?)\/( ?)(\d+)/;
	function write_num_f1(r/*:Array<string>*/, aval/*:number*/, sign/*:string*/)/*:string*/ {
		var den = parseInt(r[4],10), rr = Math.round(aval * den), base = Math.floor(rr/den);
		var myn = (rr - base*den), myd = den;
		return sign + (base === 0 ? "" : ""+base) + " " + (myn === 0 ? fill(" ", r[1].length + 1 + r[4].length) : pad_(myn,r[1].length) + r[2] + "/" + r[3] + pad0(myd,r[4].length));
	}
	function write_num_f2(r/*:Array<string>*/, aval/*:number*/, sign/*:string*/)/*:string*/ {
		return sign + (aval === 0 ? "" : ""+aval) + fill(" ", r[1].length + 2 + r[4].length);
	}
	var dec1 = /^#*0*\.([0#]+)/;
	var closeparen = /\).*[0#]/;
	var phone = /\(###\) ###\\?-####/;
	function hashq(str/*:string*/)/*:string*/ {
		var o = "", cc;
		for(var i = 0; i != str.length; ++i) switch((cc=str.charCodeAt(i))) {
			case 35: break;
			case 63: o+= " "; break;
			case 48: o+= "0"; break;
			default: o+= String.fromCharCode(cc);
		}
		return o;
	}
	function rnd(val/*:number*/, d/*:number*/)/*:string*/ { var dd = Math.pow(10,d); return ""+(Math.round(val * dd)/dd); }
	function dec(val/*:number*/, d/*:number*/)/*:number*/ {
		var _frac = val - Math.floor(val), dd = Math.pow(10,d);
		if (d < ('' + Math.round(_frac * dd)).length) return 0;
		return Math.round(_frac * dd);
	}
	function carry(val/*:number*/, d/*:number*/)/*:number*/ {
		if (d < ('' + Math.round((val-Math.floor(val))*Math.pow(10,d))).length) {
			return 1;
		}
		return 0;
	}
	function flr(val/*:number*/)/*:string*/ {
		if(val < 2147483647 && val > -2147483648) return ""+(val >= 0 ? (val|0) : (val-1|0));
		return ""+Math.floor(val);
	}
	function write_num_flt(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/ {
		if(type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
			var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
			if(val >= 0) return write_num_flt('n', ffmt, val);
			return '(' + write_num_flt('n', ffmt, -val) + ')';
		}
		if(fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm(type, fmt, val);
		if(fmt.indexOf('%') !== -1) return write_num_pct(type, fmt, val);
		if(fmt.indexOf('E') !== -1) return write_num_exp(fmt, val);
		if(fmt.charCodeAt(0) === 36) return "$"+write_num_flt(type,fmt.substr(fmt.charAt(1)==' '?2:1),val);
		var o;
		var r/*:?Array<string>*/, ri, ff, aval = Math.abs(val), sign = val < 0 ? "-" : "";
		if(fmt.match(/^00+$/)) return sign + pad0r(aval,fmt.length);
		if(fmt.match(/^[#?]+$/)) {
			o = pad0r(val,0); if(o === "0") o = "";
			return o.length > fmt.length ? o : hashq(fmt.substr(0,fmt.length-o.length)) + o;
		}
		if((r = fmt.match(frac1))) return write_num_f1(r, aval, sign);
		if(fmt.match(/^#+0+$/)) return sign + pad0r(aval,fmt.length - fmt.indexOf("0"));
		if((r = fmt.match(dec1))) {
			o = rnd(val, r[1].length).replace(/^([^\.]+)$/,"$1."+hashq(r[1])).replace(/\.$/,"."+hashq(r[1])).replace(/\.(\d*)$/,function($$, $1) { return "." + $1 + fill("0", hashq(/*::(*/r/*::||[""])*/[1]).length-$1.length); });
			return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./,".");
		}
		fmt = fmt.replace(/^#+([0.])/, "$1");
		if((r = fmt.match(/^(0*)\.(#*)$/))) {
			return sign + rnd(aval, r[2].length).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^(-?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
		}
		if((r = fmt.match(/^#{1,3},##0(\.?)$/))) return sign + commaify(pad0r(aval,0));
		if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
			return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify(""+(Math.floor(val) + carry(val, r[1].length))) + "." + pad0(dec(val, r[1].length),r[1].length);
		}
		if((r = fmt.match(/^#,#*,#0/))) return write_num_flt(type,fmt.replace(/^#,#*,/,""),val);
		if((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
			o = _strrev(write_num_flt(type, fmt.replace(/[\\-]/g,""), val));
			ri = 0;
			return _strrev(_strrev(fmt.replace(/\\/g,"")).replace(/[0#]/g,function(x){return ri<o.length?o.charAt(ri++):x==='0'?'0':"";}));
		}
		if(fmt.match(phone)) {
			o = write_num_flt(type, "##########", val);
			return "(" + o.substr(0,3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
		}
		var oa = "";
		if((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
			ri = Math.min(/*::String(*/r[4]/*::)*/.length,7);
			ff = SSF_frac(aval, Math.pow(10,ri)-1, false);
			o = "" + sign;
			oa = write_num("n", /*::String(*/r[1]/*::)*/, ff[1]);
			if(oa.charAt(oa.length-1) == " ") oa = oa.substr(0,oa.length-1) + "0";
			o += oa + /*::String(*/r[2]/*::)*/ + "/" + /*::String(*/r[3]/*::)*/;
			oa = rpad_(ff[2],ri);
			if(oa.length < r[4].length) oa = hashq(r[4].substr(r[4].length-oa.length)) + oa;
			o += oa;
			return o;
		}
		if((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
			ri = Math.min(Math.max(r[1].length, r[4].length),7);
			ff = SSF_frac(aval, Math.pow(10,ri)-1, true);
			return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad_(ff[1],ri) + r[2] + "/" + r[3] + rpad_(ff[2],ri): fill(" ", 2*ri+1 + r[2].length + r[3].length));
		}
		if((r = fmt.match(/^[#0?]+$/))) {
			o = pad0r(val, 0);
			if(fmt.length <= o.length) return o;
			return hashq(fmt.substr(0,fmt.length-o.length)) + o;
		}
		if((r = fmt.match(/^([#0?]+)\.([#0]+)$/))) {
			o = "" + val.toFixed(Math.min(r[2].length,10)).replace(/([^0])0+$/,"$1");
			ri = o.indexOf(".");
			var lres = fmt.indexOf(".") - ri, rres = fmt.length - o.length - lres;
			return hashq(fmt.substr(0,lres) + o + fmt.substr(fmt.length-rres));
		}
		if((r = fmt.match(/^00,000\.([#0]*0)$/))) {
			ri = dec(val, r[1].length);
			return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify(flr(val)).replace(/^\d,\d{3}$/,"0$&").replace(/^\d*$/,function($$) { return "00," + ($$.length < 3 ? pad0(0,3-$$.length) : "") + $$; }) + "." + pad0(ri,r[1].length);
		}
		switch(fmt) {
			case "###,##0.00": return write_num_flt(type, "#,##0.00", val);
			case "###,###":
			case "##,###":
			case "#,###": var x = commaify(pad0r(aval,0)); return x !== "0" ? sign + x : "";
			case "###,###.00": return write_num_flt(type, "###,##0.00",val).replace(/^0\./,".");
			case "#,###.00": return write_num_flt(type, "#,##0.00",val).replace(/^0\./,".");
		}
		throw new Error("unsupported format |" + fmt + "|");
	}
	function write_num_cm2(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/{
		var idx = fmt.length - 1;
		while(fmt.charCodeAt(idx-1) === 44) --idx;
		return write_num(type, fmt.substr(0,idx), val / Math.pow(10,3*(fmt.length-idx)));
	}
	function write_num_pct2(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/{
		var sfmt = fmt.replace(pct1,""), mul = fmt.length - sfmt.length;
		return write_num(type, sfmt, val * Math.pow(10,2*mul)) + fill("%",mul);
	}
	function write_num_exp2(fmt/*:string*/, val/*:number*/)/*:string*/{
		var o/*:string*/;
		var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
		if(fmt.match(/^#+0.0E\+0$/)) {
			if(val == 0) return "0.0E+0";
			else if(val < 0) return "-" + write_num_exp2(fmt, -val);
			var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
			var ee = Math.floor(Math.log(val)*Math.LOG10E)%period;
			if(ee < 0) ee += period;
			o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
			if(!o.match(/[Ee]/)) {
				var fakee = Math.floor(Math.log(val)*Math.LOG10E);
				if(o.indexOf(".") === -1) o = o.charAt(0) + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
				else o += "E+" + (fakee - ee);
				o = o.replace(/\+-/,"-");
			}
			o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
		} else o = val.toExponential(idx);
		if(fmt.match(/E\+00$/) && o.match(/e[+-]\d$/)) o = o.substr(0,o.length-1) + "0" + o.charAt(o.length-1);
		if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
		return o.replace("e","E");
	}
	function write_num_int(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/ {
		if(type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
			var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
			if(val >= 0) return write_num_int('n', ffmt, val);
			return '(' + write_num_int('n', ffmt, -val) + ')';
		}
		if(fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm2(type, fmt, val);
		if(fmt.indexOf('%') !== -1) return write_num_pct2(type, fmt, val);
		if(fmt.indexOf('E') !== -1) return write_num_exp2(fmt, val);
		if(fmt.charCodeAt(0) === 36) return "$"+write_num_int(type,fmt.substr(fmt.charAt(1)==' '?2:1),val);
		var o;
		var r/*:?Array<string>*/, ri, ff, aval = Math.abs(val), sign = val < 0 ? "-" : "";
		if(fmt.match(/^00+$/)) return sign + pad0(aval,fmt.length);
		if(fmt.match(/^[#?]+$/)) {
			o = (""+val); if(val === 0) o = "";
			return o.length > fmt.length ? o : hashq(fmt.substr(0,fmt.length-o.length)) + o;
		}
		if((r = fmt.match(frac1))) return write_num_f2(r, aval, sign);
		if(fmt.match(/^#+0+$/)) return sign + pad0(aval,fmt.length - fmt.indexOf("0"));
		if((r = fmt.match(dec1))) {
			/*:: if(!Array.isArray(r)) throw new Error("unreachable"); */
			o = (""+val).replace(/^([^\.]+)$/,"$1."+hashq(r[1])).replace(/\.$/,"."+hashq(r[1]));
			o = o.replace(/\.(\d*)$/,function($$, $1) {
			/*:: if(!Array.isArray(r)) throw new Error("unreachable"); */
				return "." + $1 + fill("0", hashq(r[1]).length-$1.length); });
			return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./,".");
		}
		fmt = fmt.replace(/^#+([0.])/, "$1");
		if((r = fmt.match(/^(0*)\.(#*)$/))) {
			return sign + (""+aval).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^(-?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
		}
		if((r = fmt.match(/^#{1,3},##0(\.?)$/))) return sign + commaify((""+aval));
		if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
			return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify((""+val)) + "." + fill('0',r[1].length);
		}
		if((r = fmt.match(/^#,#*,#0/))) return write_num_int(type,fmt.replace(/^#,#*,/,""),val);
		if((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
			o = _strrev(write_num_int(type, fmt.replace(/[\\-]/g,""), val));
			ri = 0;
			return _strrev(_strrev(fmt.replace(/\\/g,"")).replace(/[0#]/g,function(x){return ri<o.length?o.charAt(ri++):x==='0'?'0':"";}));
		}
		if(fmt.match(phone)) {
			o = write_num_int(type, "##########", val);
			return "(" + o.substr(0,3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
		}
		var oa = "";
		if((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
			ri = Math.min(/*::String(*/r[4]/*::)*/.length,7);
			ff = SSF_frac(aval, Math.pow(10,ri)-1, false);
			o = "" + sign;
			oa = write_num("n", /*::String(*/r[1]/*::)*/, ff[1]);
			if(oa.charAt(oa.length-1) == " ") oa = oa.substr(0,oa.length-1) + "0";
			o += oa + /*::String(*/r[2]/*::)*/ + "/" + /*::String(*/r[3]/*::)*/;
			oa = rpad_(ff[2],ri);
			if(oa.length < r[4].length) oa = hashq(r[4].substr(r[4].length-oa.length)) + oa;
			o += oa;
			return o;
		}
		if((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
			ri = Math.min(Math.max(r[1].length, r[4].length),7);
			ff = SSF_frac(aval, Math.pow(10,ri)-1, true);
			return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad_(ff[1],ri) + r[2] + "/" + r[3] + rpad_(ff[2],ri): fill(" ", 2*ri+1 + r[2].length + r[3].length));
		}
		if((r = fmt.match(/^[#0?]+$/))) {
			o = "" + val;
			if(fmt.length <= o.length) return o;
			return hashq(fmt.substr(0,fmt.length-o.length)) + o;
		}
		if((r = fmt.match(/^([#0]+)\.([#0]+)$/))) {
			o = "" + val.toFixed(Math.min(r[2].length,10)).replace(/([^0])0+$/,"$1");
			ri = o.indexOf(".");
			var lres = fmt.indexOf(".") - ri, rres = fmt.length - o.length - lres;
			return hashq(fmt.substr(0,lres) + o + fmt.substr(fmt.length-rres));
		}
		if((r = fmt.match(/^00,000\.([#0]*0)$/))) {
			return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify(""+val).replace(/^\d,\d{3}$/,"0$&").replace(/^\d*$/,function($$) { return "00," + ($$.length < 3 ? pad0(0,3-$$.length) : "") + $$; }) + "." + pad0(0,r[1].length);
		}
		switch(fmt) {
			case "###,###":
			case "##,###":
			case "#,###": var x = commaify(""+aval); return x !== "0" ? sign + x : "";
			default:
				if(fmt.match(/\.[0#?]*$/)) return write_num_int(type, fmt.slice(0,fmt.lastIndexOf(".")), val) + hashq(fmt.slice(fmt.lastIndexOf(".")));
		}
		throw new Error("unsupported format |" + fmt + "|");
	}
	function write_num(type/*:string*/, fmt/*:string*/, val/*:number*/)/*:string*/ {
		return (val|0) === val ? write_num_int(type, fmt, val) : write_num_flt(type, fmt, val);
	}
	function SSF_split_fmt(fmt/*:string*/)/*:Array<string>*/ {
		var out/*:Array<string>*/ = [];
		var in_str = false/*, cc*/;
		for(var i = 0, j = 0; i < fmt.length; ++i) switch((/*cc=*/fmt.charCodeAt(i))) {
			case 34: /* '"' */
				in_str = !in_str; break;
			case 95: case 42: case 92: /* '_' '*' '\\' */
				++i; break;
			case 59: /* ';' */
				out[out.length] = fmt.substr(j,i-j);
				j = i+1;
		}
		out[out.length] = fmt.substr(j);
		if(in_str === true) throw new Error("Format |" + fmt + "| unterminated string ");
		return out;
	}

	var SSF_abstime = /\[[HhMmSs\u0E0A\u0E19\u0E17]*\]/;
	function fmt_is_date(fmt/*:string*/)/*:boolean*/ {
		var i = 0, /*cc = 0,*/ c = "", o = "";
		while(i < fmt.length) {
			switch((c = fmt.charAt(i))) {
				case 'G': if(SSF_isgeneral(fmt, i)) i+= 6; i++; break;
				case '"': for(;(/*cc=*/fmt.charCodeAt(++i)) !== 34 && i < fmt.length;){/*empty*/} ++i; break;
				case '\\': i+=2; break;
				case '_': i+=2; break;
				case '@': ++i; break;
				case 'B': case 'b':
					if(fmt.charAt(i+1) === "1" || fmt.charAt(i+1) === "2") return true;
					/* falls through */
				case 'M': case 'D': case 'Y': case 'H': case 'S': case 'E':
					/* falls through */
				case 'm': case 'd': case 'y': case 'h': case 's': case 'e': case 'g': return true;
				case 'A': case 'a': case '上':
					if(fmt.substr(i, 3).toUpperCase() === "A/P") return true;
					if(fmt.substr(i, 5).toUpperCase() === "AM/PM") return true;
					if(fmt.substr(i, 5).toUpperCase() === "上午/下午") return true;
					++i; break;
				case '[':
					o = c;
					while(fmt.charAt(i++) !== ']' && i < fmt.length) o += fmt.charAt(i);
					if(o.match(SSF_abstime)) return true;
					break;
				case '.':
					/* falls through */
				case '0': case '#':
					while(i < fmt.length && ("0#?.,E+-%".indexOf(c=fmt.charAt(++i)) > -1 || (c=='\\' && fmt.charAt(i+1) == "-" && "0#".indexOf(fmt.charAt(i+2))>-1))){/* empty */}
					break;
				case '?': while(fmt.charAt(++i) === c){/* empty */} break;
				case '*': ++i; if(fmt.charAt(i) == ' ' || fmt.charAt(i) == '*') ++i; break;
				case '(': case ')': ++i; break;
				case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
					while(i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1){/* empty */} break;
				case ' ': ++i; break;
				default: ++i; break;
			}
		}
		return false;
	}

	function eval_fmt(fmt/*:string*/, v/*:any*/, opts/*:any*/, flen/*:number*/) {
		var out = [], o = "", i = 0, c = "", lst='t', dt, j, cc;
		var hr='H';
		/* Tokenize */
		while(i < fmt.length) {
			switch((c = fmt.charAt(i))) {
				case 'G': /* General */
					if(!SSF_isgeneral(fmt, i)) throw new Error('unrecognized character ' + c + ' in ' +fmt);
					out[out.length] = {t:'G', v:'General'}; i+=7; break;
				case '"': /* Literal text */
					for(o="";(cc=fmt.charCodeAt(++i)) !== 34 && i < fmt.length;) o += String.fromCharCode(cc);
					out[out.length] = {t:'t', v:o}; ++i; break;
				case '\\': var w = fmt.charAt(++i), t = (w === "(" || w === ")") ? w : 't';
					out[out.length] = {t:t, v:w}; ++i; break;
				case '_': out[out.length] = {t:'t', v:" "}; i+=2; break;
				case '@': /* Text Placeholder */
					out[out.length] = {t:'T', v:v}; ++i; break;
				case 'B': case 'b':
					if(fmt.charAt(i+1) === "1" || fmt.charAt(i+1) === "2") {
						if(dt==null) { dt=SSF_parse_date_code(v, opts, fmt.charAt(i+1) === "2"); if(dt==null) return ""; }
						out[out.length] = {t:'X', v:fmt.substr(i,2)}; lst = c; i+=2; break;
					}
					/* falls through */
				case 'M': case 'D': case 'Y': case 'H': case 'S': case 'E':
					c = c.toLowerCase();
					/* falls through */
				case 'm': case 'd': case 'y': case 'h': case 's': case 'e': case 'g':
					if(v < 0) return "";
					if(dt==null) { dt=SSF_parse_date_code(v, opts); if(dt==null) return ""; }
					o = c; while(++i < fmt.length && fmt.charAt(i).toLowerCase() === c) o+=c;
					if(c === 'm' && lst.toLowerCase() === 'h') c = 'M';
					if(c === 'h') c = hr;
					out[out.length] = {t:c, v:o}; lst = c; break;
				case 'A': case 'a': case '上':
					var q={t:c, v:c};
					if(dt==null) dt=SSF_parse_date_code(v, opts);
					if(fmt.substr(i, 3).toUpperCase() === "A/P") { if(dt!=null) q.v = dt.H >= 12 ? "P" : "A"; q.t = 'T'; hr='h';i+=3;}
					else if(fmt.substr(i,5).toUpperCase() === "AM/PM") { if(dt!=null) q.v = dt.H >= 12 ? "PM" : "AM"; q.t = 'T'; i+=5; hr='h'; }
					else if(fmt.substr(i,5).toUpperCase() === "上午/下午") { if(dt!=null) q.v = dt.H >= 12 ? "下午" : "上午"; q.t = 'T'; i+=5; hr='h'; }
					else { q.t = "t"; ++i; }
					if(dt==null && q.t === 'T') return "";
					out[out.length] = q; lst = c; break;
				case '[':
					o = c;
					while(fmt.charAt(i++) !== ']' && i < fmt.length) o += fmt.charAt(i);
					if(o.slice(-1) !== ']') throw 'unterminated "[" block: |' + o + '|';
					if(o.match(SSF_abstime)) {
						if(dt==null) { dt=SSF_parse_date_code(v, opts); if(dt==null) return ""; }
						out[out.length] = {t:'Z', v:o.toLowerCase()};
						lst = o.charAt(1);
					} else if(o.indexOf("$") > -1) {
						o = (o.match(/\$([^-\[\]]*)/)||[])[1]||"$";
						if(!fmt_is_date(fmt)) out[out.length] = {t:'t',v:o};
					}
					break;
				/* Numbers */
				case '.':
					if(dt != null) {
						o = c; while(++i < fmt.length && (c=fmt.charAt(i)) === "0") o += c;
						out[out.length] = {t:'s', v:o}; break;
					}
					/* falls through */
				case '0': case '#':
					o = c; while(++i < fmt.length && "0#?.,E+-%".indexOf(c=fmt.charAt(i)) > -1) o += c;
					out[out.length] = {t:'n', v:o}; break;
				case '?':
					o = c; while(fmt.charAt(++i) === c) o+=c;
					out[out.length] = {t:c, v:o}; lst = c; break;
				case '*': ++i; if(fmt.charAt(i) == ' ' || fmt.charAt(i) == '*') ++i; break; // **
				case '(': case ')': out[out.length] = {t:(flen===1?'t':c), v:c}; ++i; break;
				case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
					o = c; while(i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1) o+=fmt.charAt(i);
					out[out.length] = {t:'D', v:o}; break;
				case ' ': out[out.length] = {t:c, v:c}; ++i; break;
				case '$': out[out.length] = {t:'t', v:'$'}; ++i; break;
				default:
					if(",$-+/():!^&'~{}<>=€acfijklopqrtuvwxzP".indexOf(c) === -1) throw new Error('unrecognized character ' + c + ' in ' + fmt);
					out[out.length] = {t:'t', v:c}; ++i; break;
			}
		}

		/* Scan for date/time parts */
		var bt = 0, ss0 = 0, ssm;
		for(i=out.length-1, lst='t'; i >= 0; --i) {
			switch(out[i].t) {
				case 'h': case 'H': out[i].t = hr; lst='h'; if(bt < 1) bt = 1; break;
				case 's':
					if((ssm=out[i].v.match(/\.0+$/))) ss0=Math.max(ss0,ssm[0].length-1);
					if(bt < 3) bt = 3;
				/* falls through */
				case 'd': case 'y': case 'M': case 'e': lst=out[i].t; break;
				case 'm': if(lst === 's') { out[i].t = 'M'; if(bt < 2) bt = 2; } break;
				case 'X': /*if(out[i].v === "B2");*/
					break;
				case 'Z':
					if(bt < 1 && out[i].v.match(/[Hh]/)) bt = 1;
					if(bt < 2 && out[i].v.match(/[Mm]/)) bt = 2;
					if(bt < 3 && out[i].v.match(/[Ss]/)) bt = 3;
			}
		}
		/* time rounding depends on presence of minute / second / usec fields */
		switch(bt) {
			case 0: break;
			case 1:
				/*::if(!dt) break;*/
				if(dt.u >= 0.5) { dt.u = 0; ++dt.S; }
				if(dt.S >=  60) { dt.S = 0; ++dt.M; }
				if(dt.M >=  60) { dt.M = 0; ++dt.H; }
				break;
			case 2:
				/*::if(!dt) break;*/
				if(dt.u >= 0.5) { dt.u = 0; ++dt.S; }
				if(dt.S >=  60) { dt.S = 0; ++dt.M; }
				break;
		}

		/* replace fields */
		var nstr = "", jj;
		for(i=0; i < out.length; ++i) {
			switch(out[i].t) {
				case 't': case 'T': case ' ': case 'D': break;
				case 'X': out[i].v = ""; out[i].t = ";"; break;
				case 'd': case 'm': case 'y': case 'h': case 'H': case 'M': case 's': case 'e': case 'b': case 'Z':
					/*::if(!dt) throw "unreachable"; */
					out[i].v = SSF_write_date(out[i].t.charCodeAt(0), out[i].v, dt, ss0);
					out[i].t = 't'; break;
				case 'n': case '?':
					jj = i+1;
					while(out[jj] != null && (
						(c=out[jj].t) === "?" || c === "D" ||
						((c === " " || c === "t") && out[jj+1] != null && (out[jj+1].t === '?' || out[jj+1].t === "t" && out[jj+1].v === '/')) ||
						(out[i].t === '(' && (c === ' ' || c === 'n' || c === ')')) ||
						(c === 't' && (out[jj].v === '/' || out[jj].v === ' ' && out[jj+1] != null && out[jj+1].t == '?'))
					)) {
						out[i].v += out[jj].v;
						out[jj] = {v:"", t:";"}; ++jj;
					}
					nstr += out[i].v;
					i = jj-1; break;
				case 'G': out[i].t = 't'; out[i].v = SSF_general(v,opts); break;
			}
		}
		var vv = "", myv, ostr;
		if(nstr.length > 0) {
			if(nstr.charCodeAt(0) == 40) /* '(' */ {
				myv = (v<0&&nstr.charCodeAt(0) === 45 ? -v : v);
				ostr = write_num('n', nstr, myv);
			} else {
				myv = (v<0 && flen > 1 ? -v : v);
				ostr = write_num('n', nstr, myv);
				if(myv < 0 && out[0] && out[0].t == 't') {
					ostr = ostr.substr(1);
					out[0].v = "-" + out[0].v;
				}
			}
			jj=ostr.length-1;
			var decpt = out.length;
			for(i=0; i < out.length; ++i) if(out[i] != null && out[i].t != 't' && out[i].v.indexOf(".") > -1) { decpt = i; break; }
			var lasti=out.length;
			if(decpt === out.length && ostr.indexOf("E") === -1) {
				for(i=out.length-1; i>= 0;--i) {
					if(out[i] == null || 'n?'.indexOf(out[i].t) === -1) continue;
					if(jj>=out[i].v.length-1) { jj -= out[i].v.length; out[i].v = ostr.substr(jj+1, out[i].v.length); }
					else if(jj < 0) out[i].v = "";
					else { out[i].v = ostr.substr(0, jj+1); jj = -1; }
					out[i].t = 't';
					lasti = i;
				}
				if(jj>=0 && lasti<out.length) out[lasti].v = ostr.substr(0,jj+1) + out[lasti].v;
			}
			else if(decpt !== out.length && ostr.indexOf("E") === -1) {
				jj = ostr.indexOf(".")-1;
				for(i=decpt; i>= 0; --i) {
					if(out[i] == null || 'n?'.indexOf(out[i].t) === -1) continue;
					j=out[i].v.indexOf(".")>-1&&i===decpt?out[i].v.indexOf(".")-1:out[i].v.length-1;
					vv = out[i].v.substr(j+1);
					for(; j>=0; --j) {
						if(jj>=0 && (out[i].v.charAt(j) === "0" || out[i].v.charAt(j) === "#")) vv = ostr.charAt(jj--) + vv;
					}
					out[i].v = vv;
					out[i].t = 't';
					lasti = i;
				}
				if(jj>=0 && lasti<out.length) out[lasti].v = ostr.substr(0,jj+1) + out[lasti].v;
				jj = ostr.indexOf(".")+1;
				for(i=decpt; i<out.length; ++i) {
					if(out[i] == null || ('n?('.indexOf(out[i].t) === -1 && i !== decpt)) continue;
					j=out[i].v.indexOf(".")>-1&&i===decpt?out[i].v.indexOf(".")+1:0;
					vv = out[i].v.substr(0,j);
					for(; j<out[i].v.length; ++j) {
						if(jj<ostr.length) vv += ostr.charAt(jj++);
					}
					out[i].v = vv;
					out[i].t = 't';
					lasti = i;
				}
			}
		}
		for(i=0; i<out.length; ++i) if(out[i] != null && 'n?'.indexOf(out[i].t)>-1) {
			myv = (flen >1 && v < 0 && i>0 && out[i-1].v === "-" ? -v:v);
			out[i].v = write_num(out[i].t, out[i].v, myv);
			out[i].t = 't';
		}
		var retval = "";
		for(i=0; i !== out.length; ++i) if(out[i] != null) retval += out[i].v;
		return retval;
	}

	var cfregex2 = /\[(=|>[=]?|<[>=]?)(-?\d+(?:\.\d*)?)\]/;
	function chkcond(v, rr) {
		if(rr == null) return false;
		var thresh = parseFloat(rr[2]);
		switch(rr[1]) {
			case "=":  if(v == thresh) return true; break;
			case ">":  if(v >  thresh) return true; break;
			case "<":  if(v <  thresh) return true; break;
			case "<>": if(v != thresh) return true; break;
			case ">=": if(v >= thresh) return true; break;
			case "<=": if(v <= thresh) return true; break;
		}
		return false;
	}
	function choose_fmt(f/*:string*/, v/*:any*/) {
		var fmt = SSF_split_fmt(f);
		var l = fmt.length, lat = fmt[l-1].indexOf("@");
		if(l<4 && lat>-1) --l;
		if(fmt.length > 4) throw new Error("cannot find right format for |" + fmt.join("|") + "|");
		if(typeof v !== "number") return [4, fmt.length === 4 || lat>-1?fmt[fmt.length-1]:"@"];
		switch(fmt.length) {
			case 1: fmt = lat>-1 ? ["General", "General", "General", fmt[0]] : [fmt[0], fmt[0], fmt[0], "@"]; break;
			case 2: fmt = lat>-1 ? [fmt[0], fmt[0], fmt[0], fmt[1]] : [fmt[0], fmt[1], fmt[0], "@"]; break;
			case 3: fmt = lat>-1 ? [fmt[0], fmt[1], fmt[0], fmt[2]] : [fmt[0], fmt[1], fmt[2], "@"]; break;
		}
		var ff = v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2];
		if(fmt[0].indexOf("[") === -1 && fmt[1].indexOf("[") === -1) return [l, ff];
		if(fmt[0].match(/\[[=<>]/) != null || fmt[1].match(/\[[=<>]/) != null) {
			var m1 = fmt[0].match(cfregex2);
			var m2 = fmt[1].match(cfregex2);
			return chkcond(v, m1) ? [l, fmt[0]] : chkcond(v, m2) ? [l, fmt[1]] : [l, fmt[m1 != null && m2 != null ? 2 : 1]];
		}
		return [l, ff];
	}
	function SSF_format(fmt/*:string|number*/,v/*:any*/,o/*:?any*/) {
		if(o == null) o = {};
		var sfmt = "";
		switch(typeof fmt) {
			case "string":
				if(fmt == "m/d/yy" && o.dateNF) sfmt = o.dateNF;
				else sfmt = fmt;
				break;
			case "number":
				if(fmt == 14 && o.dateNF) sfmt = o.dateNF;
				else sfmt = (o.table != null ? (o.table/*:any*/) : table_fmt)[fmt];
				if(sfmt == null) sfmt = (o.table && o.table[SSF_default_map[fmt]]) || table_fmt[SSF_default_map[fmt]];
				if(sfmt == null) sfmt = SSF_default_str[fmt] || "General";
				break;
		}
		if(SSF_isgeneral(sfmt,0)) return SSF_general(v, o);
		if(v instanceof Date) v = datenum_local(v, o.date1904);
		var f = choose_fmt(sfmt, v);
		if(SSF_isgeneral(f[1])) return SSF_general(v, o);
		if(v === true) v = "TRUE"; else if(v === false) v = "FALSE";
		else if(v === "" || v == null) return "";
		return eval_fmt(f[1], v, o, f[0]);
	}
	function SSF_load(fmt/*:string*/, idx/*:?number*/)/*:number*/ {
		if(typeof idx != 'number') {
			idx = +idx || -1;
	/*::if(typeof idx != 'number') return 0x188; */
			for(var i = 0; i < 0x0188; ++i) {
	/*::if(typeof idx != 'number') return 0x188; */
				if(table_fmt[i] == undefined) { if(idx < 0) idx = i; continue; }
				if(table_fmt[i] == fmt) { idx = i; break; }
			}
	/*::if(typeof idx != 'number') return 0x188; */
			if(idx < 0) idx = 0x187;
		}
	/*::if(typeof idx != 'number') return 0x188; */
		table_fmt[idx] = fmt;
		return idx;
	}
	function SSF_load_table(tbl/*:SSFTable*/)/*:void*/ {
		for(var i=0; i!=0x0188; ++i)
			if(tbl[i] !== undefined) SSF_load(tbl[i], i);
	}

	function make_ssf() {
		table_fmt = SSF_init_table();
	}

	/*::
	declare var ReadShift:any;
	declare var CheckField:any;
	declare var prep_blob:any;
	declare var __readUInt32LE:any;
	declare var __readInt32LE:any;
	declare var __toBuffer:any;
	declare var __utf16le:any;
	declare var bconcat:any;
	declare var s2a:any;
	declare var chr0:any;
	declare var chr1:any;
	declare var has_buf:boolean;
	declare var new_buf:any;
	declare var new_raw_buf:any;
	declare var new_unsafe_buf:any;
	declare var Buffer_from:any;
	*/
	/* cfb.js (C) 2013-present SheetJS -- http://sheetjs.com */
	/* vim: set ts=2: */
	/*jshint eqnull:true */
	/*exported CFB */
	/*global Uint8Array:false, Uint16Array:false */

	/*::
	type SectorEntry = {
		name?:string;
		nodes?:Array<number>;
		data:RawBytes;
	};
	type SectorList = {
		[k:string|number]:SectorEntry;
		name:?string;
		fat_addrs:Array<number>;
		ssz:number;
	}
	type CFBFiles = {[n:string]:CFBEntry};
	*/
	/* crc32.js (C) 2014-present SheetJS -- http://sheetjs.com */
	/* vim: set ts=2: */
	/*exported CRC32 */
	var CRC32 = /*#__PURE__*/(function() {
	var CRC32 = {};
	CRC32.version = '1.2.0';
	/* see perf/crc32table.js */
	/*global Int32Array */
	function signed_crc_table()/*:any*/ {
		var c = 0, table/*:Array<number>*/ = new Array(256);

		for(var n =0; n != 256; ++n){
			c = n;
			c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
			c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
			c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
			c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
			c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
			c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
			c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
			c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
			table[n] = c;
		}

		return typeof Int32Array !== 'undefined' ? new Int32Array(table) : table;
	}

	var T0 = signed_crc_table();
	function slice_by_16_tables(T) {
		var c = 0, v = 0, n = 0, table/*:Array<number>*/ = typeof Int32Array !== 'undefined' ? new Int32Array(4096) : new Array(4096) ;

		for(n = 0; n != 256; ++n) table[n] = T[n];
		for(n = 0; n != 256; ++n) {
			v = T[n];
			for(c = 256 + n; c < 4096; c += 256) v = table[c] = (v >>> 8) ^ T[v & 0xFF];
		}
		var out = [];
		for(n = 1; n != 16; ++n) out[n - 1] = typeof Int32Array !== 'undefined' ? table.subarray(n * 256, n * 256 + 256) : table.slice(n * 256, n * 256 + 256);
		return out;
	}
	var TT = slice_by_16_tables(T0);
	var T1 = TT[0],  T2 = TT[1],  T3 = TT[2],  T4 = TT[3],  T5 = TT[4];
	var T6 = TT[5],  T7 = TT[6],  T8 = TT[7],  T9 = TT[8],  Ta = TT[9];
	var Tb = TT[10], Tc = TT[11], Td = TT[12], Te = TT[13], Tf = TT[14];
	function crc32_bstr(bstr/*:string*/, seed/*:number*/)/*:number*/ {
		var C = seed/*:: ? 0 : 0 */ ^ -1;
		for(var i = 0, L = bstr.length; i < L;) C = (C>>>8) ^ T0[(C^bstr.charCodeAt(i++))&0xFF];
		return ~C;
	}

	function crc32_buf(B/*:Uint8Array|Array<number>*/, seed/*:number*/)/*:number*/ {
		var C = seed/*:: ? 0 : 0 */ ^ -1, L = B.length - 15, i = 0;
		for(; i < L;) C =
			Tf[B[i++] ^ (C & 255)] ^
			Te[B[i++] ^ ((C >> 8) & 255)] ^
			Td[B[i++] ^ ((C >> 16) & 255)] ^
			Tc[B[i++] ^ (C >>> 24)] ^
			Tb[B[i++]] ^ Ta[B[i++]] ^ T9[B[i++]] ^ T8[B[i++]] ^
			T7[B[i++]] ^ T6[B[i++]] ^ T5[B[i++]] ^ T4[B[i++]] ^
			T3[B[i++]] ^ T2[B[i++]] ^ T1[B[i++]] ^ T0[B[i++]];
		L += 15;
		while(i < L) C = (C>>>8) ^ T0[(C^B[i++])&0xFF];
		return ~C;
	}

	function crc32_str(str/*:string*/, seed/*:number*/)/*:number*/ {
		var C = seed ^ -1;
		for(var i = 0, L = str.length, c = 0, d = 0; i < L;) {
			c = str.charCodeAt(i++);
			if(c < 0x80) {
				C = (C>>>8) ^ T0[(C^c)&0xFF];
			} else if(c < 0x800) {
				C = (C>>>8) ^ T0[(C ^ (192|((c>>6)&31)))&0xFF];
				C = (C>>>8) ^ T0[(C ^ (128|(c&63)))&0xFF];
			} else if(c >= 0xD800 && c < 0xE000) {
				c = (c&1023)+64; d = str.charCodeAt(i++)&1023;
				C = (C>>>8) ^ T0[(C ^ (240|((c>>8)&7)))&0xFF];
				C = (C>>>8) ^ T0[(C ^ (128|((c>>2)&63)))&0xFF];
				C = (C>>>8) ^ T0[(C ^ (128|((d>>6)&15)|((c&3)<<4)))&0xFF];
				C = (C>>>8) ^ T0[(C ^ (128|(d&63)))&0xFF];
			} else {
				C = (C>>>8) ^ T0[(C ^ (224|((c>>12)&15)))&0xFF];
				C = (C>>>8) ^ T0[(C ^ (128|((c>>6)&63)))&0xFF];
				C = (C>>>8) ^ T0[(C ^ (128|(c&63)))&0xFF];
			}
		}
		return ~C;
	}
	CRC32.table = T0;
	CRC32.bstr = crc32_bstr;
	CRC32.buf = crc32_buf;
	CRC32.str = crc32_str;
	return CRC32;
	})();
	/* [MS-CFB] v20171201 */
	var CFB = /*#__PURE__*/(function _CFB(){
	var exports = {};
	exports.version = '1.2.1';
	/* [MS-CFB] 2.6.4 */
	function namecmp(l/*:string*/, r/*:string*/)/*:number*/ {
		var L = l.split("/"), R = r.split("/");
		for(var i = 0, c = 0, Z = Math.min(L.length, R.length); i < Z; ++i) {
			if((c = L[i].length - R[i].length)) return c;
			if(L[i] != R[i]) return L[i] < R[i] ? -1 : 1;
		}
		return L.length - R.length;
	}
	function dirname(p/*:string*/)/*:string*/ {
		if(p.charAt(p.length - 1) == "/") return (p.slice(0,-1).indexOf("/") === -1) ? p : dirname(p.slice(0, -1));
		var c = p.lastIndexOf("/");
		return (c === -1) ? p : p.slice(0, c+1);
	}

	function filename(p/*:string*/)/*:string*/ {
		if(p.charAt(p.length - 1) == "/") return filename(p.slice(0, -1));
		var c = p.lastIndexOf("/");
		return (c === -1) ? p : p.slice(c+1);
	}
	/* -------------------------------------------------------------------------- */
	/* DOS Date format:
	   high|YYYYYYYm.mmmddddd.HHHHHMMM.MMMSSSSS|low
	   add 1980 to stored year
	   stored second should be doubled
	*/

	/* write JS date to buf as a DOS date */
	function write_dos_date(buf/*:CFBlob*/, date/*:Date|string*/) {
		if(typeof date === "string") date = new Date(date);
		var hms/*:number*/ = date.getHours();
		hms = hms << 6 | date.getMinutes();
		hms = hms << 5 | (date.getSeconds()>>>1);
		buf.write_shift(2, hms);
		var ymd/*:number*/ = (date.getFullYear() - 1980);
		ymd = ymd << 4 | (date.getMonth()+1);
		ymd = ymd << 5 | date.getDate();
		buf.write_shift(2, ymd);
	}

	/* read four bytes from buf and interpret as a DOS date */
	function parse_dos_date(buf/*:CFBlob*/)/*:Date*/ {
		var hms = buf.read_shift(2) & 0xFFFF;
		var ymd = buf.read_shift(2) & 0xFFFF;
		var val = new Date();
		var d = ymd & 0x1F; ymd >>>= 5;
		var m = ymd & 0x0F; ymd >>>= 4;
		val.setMilliseconds(0);
		val.setFullYear(ymd + 1980);
		val.setMonth(m-1);
		val.setDate(d);
		var S = hms & 0x1F; hms >>>= 5;
		var M = hms & 0x3F; hms >>>= 6;
		val.setHours(hms);
		val.setMinutes(M);
		val.setSeconds(S<<1);
		return val;
	}
	function parse_extra_field(blob/*:CFBlob*/)/*:any*/ {
		prep_blob(blob, 0);
		var o = /*::(*/{}/*:: :any)*/;
		var flags = 0;
		while(blob.l <= blob.length - 4) {
			var type = blob.read_shift(2);
			var sz = blob.read_shift(2), tgt = blob.l + sz;
			var p = {};
			switch(type) {
				/* UNIX-style Timestamps */
				case 0x5455: {
					flags = blob.read_shift(1);
					if(flags & 1) p.mtime = blob.read_shift(4);
					/* for some reason, CD flag corresponds to LFH */
					if(sz > 5) {
						if(flags & 2) p.atime = blob.read_shift(4);
						if(flags & 4) p.ctime = blob.read_shift(4);
					}
					if(p.mtime) p.mt = new Date(p.mtime*1000);
				}
				break;
			}
			blob.l = tgt;
			o[type] = p;
		}
		return o;
	}
	var fs/*:: = require('fs'); */;
	function get_fs() { return fs || (fs = {}); }
	function parse(file/*:RawBytes*/, options/*:CFBReadOpts*/)/*:CFBContainer*/ {
	if(file[0] == 0x50 && file[1] == 0x4b) return parse_zip(file, options);
	if((file[0] | 0x20) == 0x6d && (file[1]|0x20) == 0x69) return parse_mad(file, options);
	if(file.length < 512) throw new Error("CFB file size " + file.length + " < 512");
	var mver = 3;
	var ssz = 512;
	var nmfs = 0; // number of mini FAT sectors
	var difat_sec_cnt = 0;
	var dir_start = 0;
	var minifat_start = 0;
	var difat_start = 0;

	var fat_addrs/*:Array<number>*/ = []; // locations of FAT sectors

	/* [MS-CFB] 2.2 Compound File Header */
	var blob/*:CFBlob*/ = /*::(*/file.slice(0,512)/*:: :any)*/;
	prep_blob(blob, 0);

	/* major version */
	var mv = check_get_mver(blob);
	mver = mv[0];
	switch(mver) {
		case 3: ssz = 512; break; case 4: ssz = 4096; break;
		case 0: if(mv[1] == 0) return parse_zip(file, options);
		/* falls through */
		default: throw new Error("Major Version: Expected 3 or 4 saw " + mver);
	}

	/* reprocess header */
	if(ssz !== 512) { blob = /*::(*/file.slice(0,ssz)/*:: :any)*/; prep_blob(blob, 28 /* blob.l */); }
	/* Save header for final object */
	var header/*:RawBytes*/ = file.slice(0,ssz);

	check_shifts(blob, mver);

	// Number of Directory Sectors
	var dir_cnt/*:number*/ = blob.read_shift(4, 'i');
	if(mver === 3 && dir_cnt !== 0) throw new Error('# Directory Sectors: Expected 0 saw ' + dir_cnt);

	// Number of FAT Sectors
	blob.l += 4;

	// First Directory Sector Location
	dir_start = blob.read_shift(4, 'i');

	// Transaction Signature
	blob.l += 4;

	// Mini Stream Cutoff Size
	blob.chk('00100000', 'Mini Stream Cutoff Size: ');

	// First Mini FAT Sector Location
	minifat_start = blob.read_shift(4, 'i');

	// Number of Mini FAT Sectors
	nmfs = blob.read_shift(4, 'i');

	// First DIFAT sector location
	difat_start = blob.read_shift(4, 'i');

	// Number of DIFAT Sectors
	difat_sec_cnt = blob.read_shift(4, 'i');

	// Grab FAT Sector Locations
	for(var q = -1, j = 0; j < 109; ++j) { /* 109 = (512 - blob.l)>>>2; */
		q = blob.read_shift(4, 'i');
		if(q<0) break;
		fat_addrs[j] = q;
	}

	/** Break the file up into sectors */
	var sectors/*:Array<RawBytes>*/ = sectorify(file, ssz);

	sleuth_fat(difat_start, difat_sec_cnt, sectors, ssz, fat_addrs);

	/** Chains */
	var sector_list/*:SectorList*/ = make_sector_list(sectors, dir_start, fat_addrs, ssz);

	sector_list[dir_start].name = "!Directory";
	if(nmfs > 0 && minifat_start !== ENDOFCHAIN) sector_list[minifat_start].name = "!MiniFAT";
	sector_list[fat_addrs[0]].name = "!FAT";
	sector_list.fat_addrs = fat_addrs;
	sector_list.ssz = ssz;

	/* [MS-CFB] 2.6.1 Compound File Directory Entry */
	var files/*:CFBFiles*/ = {}, Paths/*:Array<string>*/ = [], FileIndex/*:CFBFileIndex*/ = [], FullPaths/*:Array<string>*/ = [];
	read_directory(dir_start, sector_list, sectors, Paths, nmfs, files, FileIndex, minifat_start);

	build_full_paths(FileIndex, FullPaths, Paths);
	Paths.shift();

	var o = {
		FileIndex: FileIndex,
		FullPaths: FullPaths
	};

	// $FlowIgnore
	if(options && options.raw) o.raw = {header: header, sectors: sectors};
	return o;
	} // parse

	/* [MS-CFB] 2.2 Compound File Header -- read up to major version */
	function check_get_mver(blob/*:CFBlob*/)/*:[number, number]*/ {
		if(blob[blob.l] == 0x50 && blob[blob.l + 1] == 0x4b) return [0, 0];
		// header signature 8
		blob.chk(HEADER_SIGNATURE, 'Header Signature: ');

		// clsid 16
		//blob.chk(HEADER_CLSID, 'CLSID: ');
		blob.l += 16;

		// minor version 2
		var mver/*:number*/ = blob.read_shift(2, 'u');

		return [blob.read_shift(2,'u'), mver];
	}
	function check_shifts(blob/*:CFBlob*/, mver/*:number*/)/*:void*/ {
		var shift = 0x09;

		// Byte Order
		//blob.chk('feff', 'Byte Order: '); // note: some writers put 0xffff
		blob.l += 2;

		// Sector Shift
		switch((shift = blob.read_shift(2))) {
			case 0x09: if(mver != 3) throw new Error('Sector Shift: Expected 9 saw ' + shift); break;
			case 0x0c: if(mver != 4) throw new Error('Sector Shift: Expected 12 saw ' + shift); break;
			default: throw new Error('Sector Shift: Expected 9 or 12 saw ' + shift);
		}

		// Mini Sector Shift
		blob.chk('0600', 'Mini Sector Shift: ');

		// Reserved
		blob.chk('000000000000', 'Reserved: ');
	}

	/** Break the file up into sectors */
	function sectorify(file/*:RawBytes*/, ssz/*:number*/)/*:Array<RawBytes>*/ {
		var nsectors = Math.ceil(file.length/ssz)-1;
		var sectors/*:Array<RawBytes>*/ = [];
		for(var i=1; i < nsectors; ++i) sectors[i-1] = file.slice(i*ssz,(i+1)*ssz);
		sectors[nsectors-1] = file.slice(nsectors*ssz);
		return sectors;
	}

	/* [MS-CFB] 2.6.4 Red-Black Tree */
	function build_full_paths(FI/*:CFBFileIndex*/, FP/*:Array<string>*/, Paths/*:Array<string>*/)/*:void*/ {
		var i = 0, L = 0, R = 0, C = 0, j = 0, pl = Paths.length;
		var dad/*:Array<number>*/ = [], q/*:Array<number>*/ = [];

		for(; i < pl; ++i) { dad[i]=q[i]=i; FP[i]=Paths[i]; }

		for(; j < q.length; ++j) {
			i = q[j];
			L = FI[i].L; R = FI[i].R; C = FI[i].C;
			if(dad[i] === i) {
				if(L !== -1 /*NOSTREAM*/ && dad[L] !== L) dad[i] = dad[L];
				if(R !== -1 && dad[R] !== R) dad[i] = dad[R];
			}
			if(C !== -1 /*NOSTREAM*/) dad[C] = i;
			if(L !== -1 && i != dad[i]) { dad[L] = dad[i]; if(q.lastIndexOf(L) < j) q.push(L); }
			if(R !== -1 && i != dad[i]) { dad[R] = dad[i]; if(q.lastIndexOf(R) < j) q.push(R); }
		}
		for(i=1; i < pl; ++i) if(dad[i] === i) {
			if(R !== -1 /*NOSTREAM*/ && dad[R] !== R) dad[i] = dad[R];
			else if(L !== -1 && dad[L] !== L) dad[i] = dad[L];
		}

		for(i=1; i < pl; ++i) {
			if(FI[i].type === 0 /* unknown */) continue;
			j = i;
			if(j != dad[j]) do {
				j = dad[j];
				FP[i] = FP[j] + "/" + FP[i];
			} while (j !== 0 && -1 !== dad[j] && j != dad[j]);
			dad[i] = -1;
		}

		FP[0] += "/";
		for(i=1; i < pl; ++i) {
			if(FI[i].type !== 2 /* stream */) FP[i] += "/";
		}
	}

	function get_mfat_entry(entry/*:CFBEntry*/, payload/*:RawBytes*/, mini/*:?RawBytes*/)/*:CFBlob*/ {
		var start = entry.start, size = entry.size;
		//return (payload.slice(start*MSSZ, start*MSSZ + size)/*:any*/);
		var o = [];
		var idx = start;
		while(mini && size > 0 && idx >= 0) {
			o.push(payload.slice(idx * MSSZ, idx * MSSZ + MSSZ));
			size -= MSSZ;
			idx = __readInt32LE(mini, idx * 4);
		}
		if(o.length === 0) return (new_buf(0)/*:any*/);
		return (bconcat(o).slice(0, entry.size)/*:any*/);
	}

	/** Chase down the rest of the DIFAT chain to build a comprehensive list
	    DIFAT chains by storing the next sector number as the last 32 bits */
	function sleuth_fat(idx/*:number*/, cnt/*:number*/, sectors/*:Array<RawBytes>*/, ssz/*:number*/, fat_addrs)/*:void*/ {
		var q/*:number*/ = ENDOFCHAIN;
		if(idx === ENDOFCHAIN) {
			if(cnt !== 0) throw new Error("DIFAT chain shorter than expected");
		} else if(idx !== -1 /*FREESECT*/) {
			var sector = sectors[idx], m = (ssz>>>2)-1;
			if(!sector) return;
			for(var i = 0; i < m; ++i) {
				if((q = __readInt32LE(sector,i*4)) === ENDOFCHAIN) break;
				fat_addrs.push(q);
			}
			sleuth_fat(__readInt32LE(sector,ssz-4),cnt - 1, sectors, ssz, fat_addrs);
		}
	}

	/** Follow the linked list of sectors for a given starting point */
	function get_sector_list(sectors/*:Array<RawBytes>*/, start/*:number*/, fat_addrs/*:Array<number>*/, ssz/*:number*/, chkd/*:?Array<boolean>*/)/*:SectorEntry*/ {
		var buf/*:Array<number>*/ = [], buf_chain/*:Array<any>*/ = [];
		if(!chkd) chkd = [];
		var modulus = ssz - 1, j = 0, jj = 0;
		for(j=start; j>=0;) {
			chkd[j] = true;
			buf[buf.length] = j;
			buf_chain.push(sectors[j]);
			var addr = fat_addrs[Math.floor(j*4/ssz)];
			jj = ((j*4) & modulus);
			if(ssz < 4 + jj) throw new Error("FAT boundary crossed: " + j + " 4 "+ssz);
			if(!sectors[addr]) break;
			j = __readInt32LE(sectors[addr], jj);
		}
		return {nodes: buf, data:__toBuffer([buf_chain])};
	}

	/** Chase down the sector linked lists */
	function make_sector_list(sectors/*:Array<RawBytes>*/, dir_start/*:number*/, fat_addrs/*:Array<number>*/, ssz/*:number*/)/*:SectorList*/ {
		var sl = sectors.length, sector_list/*:SectorList*/ = ([]/*:any*/);
		var chkd/*:Array<boolean>*/ = [], buf/*:Array<number>*/ = [], buf_chain/*:Array<RawBytes>*/ = [];
		var modulus = ssz - 1, i=0, j=0, k=0, jj=0;
		for(i=0; i < sl; ++i) {
			buf = ([]/*:Array<number>*/);
			k = (i + dir_start); if(k >= sl) k-=sl;
			if(chkd[k]) continue;
			buf_chain = [];
			var seen = [];
			for(j=k; j>=0;) {
				seen[j] = true;
				chkd[j] = true;
				buf[buf.length] = j;
				buf_chain.push(sectors[j]);
				var addr/*:number*/ = fat_addrs[Math.floor(j*4/ssz)];
				jj = ((j*4) & modulus);
				if(ssz < 4 + jj) throw new Error("FAT boundary crossed: " + j + " 4 "+ssz);
				if(!sectors[addr]) break;
				j = __readInt32LE(sectors[addr], jj);
				if(seen[j]) break;
			}
			sector_list[k] = ({nodes: buf, data:__toBuffer([buf_chain])}/*:SectorEntry*/);
		}
		return sector_list;
	}

	/* [MS-CFB] 2.6.1 Compound File Directory Entry */
	function read_directory(dir_start/*:number*/, sector_list/*:SectorList*/, sectors/*:Array<RawBytes>*/, Paths/*:Array<string>*/, nmfs, files, FileIndex, mini) {
		var minifat_store = 0, pl = (Paths.length?2:0);
		var sector = sector_list[dir_start].data;
		var i = 0, namelen = 0, name;
		for(; i < sector.length; i+= 128) {
			var blob/*:CFBlob*/ = /*::(*/sector.slice(i, i+128)/*:: :any)*/;
			prep_blob(blob, 64);
			namelen = blob.read_shift(2);
			name = __utf16le(blob,0,namelen-pl);
			Paths.push(name);
			var o/*:CFBEntry*/ = ({
				name:  name,
				type:  blob.read_shift(1),
				color: blob.read_shift(1),
				L:     blob.read_shift(4, 'i'),
				R:     blob.read_shift(4, 'i'),
				C:     blob.read_shift(4, 'i'),
				clsid: blob.read_shift(16),
				state: blob.read_shift(4, 'i'),
				start: 0,
				size: 0
			});
			var ctime/*:number*/ = blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2);
			if(ctime !== 0) o.ct = read_date(blob, blob.l-8);
			var mtime/*:number*/ = blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2);
			if(mtime !== 0) o.mt = read_date(blob, blob.l-8);
			o.start = blob.read_shift(4, 'i');
			o.size = blob.read_shift(4, 'i');
			if(o.size < 0 && o.start < 0) { o.size = o.type = 0; o.start = ENDOFCHAIN; o.name = ""; }
			if(o.type === 5) { /* root */
				minifat_store = o.start;
				if(nmfs > 0 && minifat_store !== ENDOFCHAIN) sector_list[minifat_store].name = "!StreamData";
				/*minifat_size = o.size;*/
			} else if(o.size >= 4096 /* MSCSZ */) {
				o.storage = 'fat';
				if(sector_list[o.start] === undefined) sector_list[o.start] = get_sector_list(sectors, o.start, sector_list.fat_addrs, sector_list.ssz);
				sector_list[o.start].name = o.name;
				o.content = (sector_list[o.start].data.slice(0,o.size)/*:any*/);
			} else {
				o.storage = 'minifat';
				if(o.size < 0) o.size = 0;
				else if(minifat_store !== ENDOFCHAIN && o.start !== ENDOFCHAIN && sector_list[minifat_store]) {
					o.content = get_mfat_entry(o, sector_list[minifat_store].data, (sector_list[mini]||{}).data);
				}
			}
			if(o.content) prep_blob(o.content, 0);
			files[name] = o;
			FileIndex.push(o);
		}
	}

	function read_date(blob/*:RawBytes|CFBlob*/, offset/*:number*/)/*:Date*/ {
		return new Date(( ( (__readUInt32LE(blob,offset+4)/1e7)*Math.pow(2,32)+__readUInt32LE(blob,offset)/1e7 ) - 11644473600)*1000);
	}

	function read_file(filename/*:string*/, options/*:CFBReadOpts*/) {
		get_fs();
		return parse(fs.readFileSync(filename), options);
	}

	function read(blob/*:RawBytes|string*/, options/*:CFBReadOpts*/) {
		var type = options && options.type;
		if(!type) {
			if(has_buf && Buffer.isBuffer(blob)) type = "buffer";
		}
		switch(type || "base64") {
			case "file": /*:: if(typeof blob !== 'string') throw "Must pass a filename when type='file'"; */return read_file(blob, options);
			case "base64": /*:: if(typeof blob !== 'string') throw "Must pass a base64-encoded binary string when type='file'"; */return parse(s2a(Base64_decode(blob)), options);
			case "binary": /*:: if(typeof blob !== 'string') throw "Must pass a binary string when type='file'"; */return parse(s2a(blob), options);
		}
		return parse(/*::typeof blob == 'string' ? new Buffer(blob, 'utf-8') : */blob, options);
	}

	function init_cfb(cfb/*:CFBContainer*/, opts/*:?any*/)/*:void*/ {
		var o = opts || {}, root = o.root || "Root Entry";
		if(!cfb.FullPaths) cfb.FullPaths = [];
		if(!cfb.FileIndex) cfb.FileIndex = [];
		if(cfb.FullPaths.length !== cfb.FileIndex.length) throw new Error("inconsistent CFB structure");
		if(cfb.FullPaths.length === 0) {
			cfb.FullPaths[0] = root + "/";
			cfb.FileIndex[0] = ({ name: root, type: 5 }/*:any*/);
		}
		if(o.CLSID) cfb.FileIndex[0].clsid = o.CLSID;
		seed_cfb(cfb);
	}
	function seed_cfb(cfb/*:CFBContainer*/)/*:void*/ {
		var nm = "\u0001Sh33tJ5";
		if(CFB.find(cfb, "/" + nm)) return;
		var p = new_buf(4); p[0] = 55; p[1] = p[3] = 50; p[2] = 54;
		cfb.FileIndex.push(({ name: nm, type: 2, content:p, size:4, L:69, R:69, C:69 }/*:any*/));
		cfb.FullPaths.push(cfb.FullPaths[0] + nm);
		rebuild_cfb(cfb);
	}
	function rebuild_cfb(cfb/*:CFBContainer*/, f/*:?boolean*/)/*:void*/ {
		init_cfb(cfb);
		var gc = false, s = false;
		for(var i = cfb.FullPaths.length - 1; i >= 0; --i) {
			var _file = cfb.FileIndex[i];
			switch(_file.type) {
				case 0:
					if(s) gc = true;
					else { cfb.FileIndex.pop(); cfb.FullPaths.pop(); }
					break;
				case 1: case 2: case 5:
					s = true;
					if(isNaN(_file.R * _file.L * _file.C)) gc = true;
					if(_file.R > -1 && _file.L > -1 && _file.R == _file.L) gc = true;
					break;
				default: gc = true; break;
			}
		}
		if(!gc && !f) return;

		var now = new Date(1987, 1, 19), j = 0;
		// Track which names exist
		var fullPaths = Object.create ? Object.create(null) : {};
		var data/*:Array<[string, CFBEntry]>*/ = [];
		for(i = 0; i < cfb.FullPaths.length; ++i) {
			fullPaths[cfb.FullPaths[i]] = true;
			if(cfb.FileIndex[i].type === 0) continue;
			data.push([cfb.FullPaths[i], cfb.FileIndex[i]]);
		}
		for(i = 0; i < data.length; ++i) {
			var dad = dirname(data[i][0]);
			s = fullPaths[dad];
			if(!s) {
				data.push([dad, ({
					name: filename(dad).replace("/",""),
					type: 1,
					clsid: HEADER_CLSID,
					ct: now, mt: now,
					content: null
				}/*:any*/)]);
				// Add name to set
				fullPaths[dad] = true;
			}
		}

		data.sort(function(x,y) { return namecmp(x[0], y[0]); });
		cfb.FullPaths = []; cfb.FileIndex = [];
		for(i = 0; i < data.length; ++i) { cfb.FullPaths[i] = data[i][0]; cfb.FileIndex[i] = data[i][1]; }
		for(i = 0; i < data.length; ++i) {
			var elt = cfb.FileIndex[i];
			var nm = cfb.FullPaths[i];

			elt.name =  filename(nm).replace("/","");
			elt.L = elt.R = elt.C = -(elt.color = 1);
			elt.size = elt.content ? elt.content.length : 0;
			elt.start = 0;
			elt.clsid = (elt.clsid || HEADER_CLSID);
			if(i === 0) {
				elt.C = data.length > 1 ? 1 : -1;
				elt.size = 0;
				elt.type = 5;
			} else if(nm.slice(-1) == "/") {
				for(j=i+1;j < data.length; ++j) if(dirname(cfb.FullPaths[j])==nm) break;
				elt.C = j >= data.length ? -1 : j;
				for(j=i+1;j < data.length; ++j) if(dirname(cfb.FullPaths[j])==dirname(nm)) break;
				elt.R = j >= data.length ? -1 : j;
				elt.type = 1;
			} else {
				if(dirname(cfb.FullPaths[i+1]||"") == dirname(nm)) elt.R = i + 1;
				elt.type = 2;
			}
		}

	}

	function _write(cfb/*:CFBContainer*/, options/*:CFBWriteOpts*/)/*:RawBytes|string*/ {
		var _opts = options || {};
		/* MAD is order-sensitive, skip rebuild and sort */
		if(_opts.fileType == 'mad') return write_mad(cfb, _opts);
		rebuild_cfb(cfb);
		switch(_opts.fileType) {
			case 'zip': return write_zip(cfb, _opts);
			//case 'mad': return write_mad(cfb, _opts);
		}
		var L = (function(cfb/*:CFBContainer*/)/*:Array<number>*/{
			var mini_size = 0, fat_size = 0;
			for(var i = 0; i < cfb.FileIndex.length; ++i) {
				var file = cfb.FileIndex[i];
				if(!file.content) continue;
				/*:: if(file.content == null) throw new Error("unreachable"); */
				var flen = file.content.length;
				if(flen > 0){
					if(flen < 0x1000) mini_size += (flen + 0x3F) >> 6;
					else fat_size += (flen + 0x01FF) >> 9;
				}
			}
			var dir_cnt = (cfb.FullPaths.length +3) >> 2;
			var mini_cnt = (mini_size + 7) >> 3;
			var mfat_cnt = (mini_size + 0x7F) >> 7;
			var fat_base = mini_cnt + fat_size + dir_cnt + mfat_cnt;
			var fat_cnt = (fat_base + 0x7F) >> 7;
			var difat_cnt = fat_cnt <= 109 ? 0 : Math.ceil((fat_cnt-109)/0x7F);
			while(((fat_base + fat_cnt + difat_cnt + 0x7F) >> 7) > fat_cnt) difat_cnt = ++fat_cnt <= 109 ? 0 : Math.ceil((fat_cnt-109)/0x7F);
			var L =  [1, difat_cnt, fat_cnt, mfat_cnt, dir_cnt, fat_size, mini_size, 0];
			cfb.FileIndex[0].size = mini_size << 6;
			L[7] = (cfb.FileIndex[0].start=L[0]+L[1]+L[2]+L[3]+L[4]+L[5])+((L[6]+7) >> 3);
			return L;
		})(cfb);
		var o = new_buf(L[7] << 9);
		var i = 0, T = 0;
		{
			for(i = 0; i < 8; ++i) o.write_shift(1, HEADER_SIG[i]);
			for(i = 0; i < 8; ++i) o.write_shift(2, 0);
			o.write_shift(2, 0x003E);
			o.write_shift(2, 0x0003);
			o.write_shift(2, 0xFFFE);
			o.write_shift(2, 0x0009);
			o.write_shift(2, 0x0006);
			for(i = 0; i < 3; ++i) o.write_shift(2, 0);
			o.write_shift(4, 0);
			o.write_shift(4, L[2]);
			o.write_shift(4, L[0] + L[1] + L[2] + L[3] - 1);
			o.write_shift(4, 0);
			o.write_shift(4, 1<<12);
			o.write_shift(4, L[3] ? L[0] + L[1] + L[2] - 1: ENDOFCHAIN);
			o.write_shift(4, L[3]);
			o.write_shift(-4, L[1] ? L[0] - 1: ENDOFCHAIN);
			o.write_shift(4, L[1]);
			for(i = 0; i < 109; ++i) o.write_shift(-4, i < L[2] ? L[1] + i : -1);
		}
		if(L[1]) {
			for(T = 0; T < L[1]; ++T) {
				for(; i < 236 + T * 127; ++i) o.write_shift(-4, i < L[2] ? L[1] + i : -1);
				o.write_shift(-4, T === L[1] - 1 ? ENDOFCHAIN : T + 1);
			}
		}
		var chainit = function(w/*:number*/)/*:void*/ {
			for(T += w; i<T-1; ++i) o.write_shift(-4, i+1);
			if(w) { ++i; o.write_shift(-4, ENDOFCHAIN); }
		};
		T = i = 0;
		for(T+=L[1]; i<T; ++i) o.write_shift(-4, consts.DIFSECT);
		for(T+=L[2]; i<T; ++i) o.write_shift(-4, consts.FATSECT);
		chainit(L[3]);
		chainit(L[4]);
		var j/*:number*/ = 0, flen/*:number*/ = 0;
		var file/*:CFBEntry*/ = cfb.FileIndex[0];
		for(; j < cfb.FileIndex.length; ++j) {
			file = cfb.FileIndex[j];
			if(!file.content) continue;
			/*:: if(file.content == null) throw new Error("unreachable"); */
			flen = file.content.length;
			if(flen < 0x1000) continue;
			file.start = T;
			chainit((flen + 0x01FF) >> 9);
		}
		chainit((L[6] + 7) >> 3);
		while(o.l & 0x1FF) o.write_shift(-4, consts.ENDOFCHAIN);
		T = i = 0;
		for(j = 0; j < cfb.FileIndex.length; ++j) {
			file = cfb.FileIndex[j];
			if(!file.content) continue;
			/*:: if(file.content == null) throw new Error("unreachable"); */
			flen = file.content.length;
			if(!flen || flen >= 0x1000) continue;
			file.start = T;
			chainit((flen + 0x3F) >> 6);
		}
		while(o.l & 0x1FF) o.write_shift(-4, consts.ENDOFCHAIN);
		for(i = 0; i < L[4]<<2; ++i) {
			var nm = cfb.FullPaths[i];
			if(!nm || nm.length === 0) {
				for(j = 0; j < 17; ++j) o.write_shift(4, 0);
				for(j = 0; j < 3; ++j) o.write_shift(4, -1);
				for(j = 0; j < 12; ++j) o.write_shift(4, 0);
				continue;
			}
			file = cfb.FileIndex[i];
			if(i === 0) file.start = file.size ? file.start - 1 : ENDOFCHAIN;
			var _nm/*:string*/ = (i === 0 && _opts.root) || file.name;
			flen = 2*(_nm.length+1);
			o.write_shift(64, _nm, "utf16le");
			o.write_shift(2, flen);
			o.write_shift(1, file.type);
			o.write_shift(1, file.color);
			o.write_shift(-4, file.L);
			o.write_shift(-4, file.R);
			o.write_shift(-4, file.C);
			if(!file.clsid) for(j = 0; j < 4; ++j) o.write_shift(4, 0);
			else o.write_shift(16, file.clsid, "hex");
			o.write_shift(4, file.state || 0);
			o.write_shift(4, 0); o.write_shift(4, 0);
			o.write_shift(4, 0); o.write_shift(4, 0);
			o.write_shift(4, file.start);
			o.write_shift(4, file.size); o.write_shift(4, 0);
		}
		for(i = 1; i < cfb.FileIndex.length; ++i) {
			file = cfb.FileIndex[i];
			/*:: if(!file.content) throw new Error("unreachable"); */
			if(file.size >= 0x1000) {
				o.l = (file.start+1) << 9;
				if (has_buf && Buffer.isBuffer(file.content)) {
					file.content.copy(o, o.l, 0, file.size);
					// o is a 0-filled Buffer so just set next offset
					o.l += (file.size + 511) & -512;
				} else {
					for(j = 0; j < file.size; ++j) o.write_shift(1, file.content[j]);
					for(; j & 0x1FF; ++j) o.write_shift(1, 0);
				}
			}
		}
		for(i = 1; i < cfb.FileIndex.length; ++i) {
			file = cfb.FileIndex[i];
			/*:: if(!file.content) throw new Error("unreachable"); */
			if(file.size > 0 && file.size < 0x1000) {
				if (has_buf && Buffer.isBuffer(file.content)) {
					file.content.copy(o, o.l, 0, file.size);
					// o is a 0-filled Buffer so just set next offset
					o.l += (file.size + 63) & -64;
				} else {
					for(j = 0; j < file.size; ++j) o.write_shift(1, file.content[j]);
					for(; j & 0x3F; ++j) o.write_shift(1, 0);
				}
			}
		}
		if (has_buf) {
			o.l = o.length;
		} else {
			// When using Buffer, already 0-filled
			while(o.l < o.length) o.write_shift(1, 0);
		}
		return o;
	}
	/* [MS-CFB] 2.6.4 (Unicode 3.0.1 case conversion) */
	function find(cfb/*:CFBContainer*/, path/*:string*/)/*:?CFBEntry*/ {
		var UCFullPaths/*:Array<string>*/ = cfb.FullPaths.map(function(x) { return x.toUpperCase(); });
		var UCPaths/*:Array<string>*/ = UCFullPaths.map(function(x) { var y = x.split("/"); return y[y.length - (x.slice(-1) == "/" ? 2 : 1)]; });
		var k/*:boolean*/ = false;
		if(path.charCodeAt(0) === 47 /* "/" */) { k = true; path = UCFullPaths[0].slice(0, -1) + path; }
		else k = path.indexOf("/") !== -1;
		var UCPath/*:string*/ = path.toUpperCase();
		var w/*:number*/ = k === true ? UCFullPaths.indexOf(UCPath) : UCPaths.indexOf(UCPath);
		if(w !== -1) return cfb.FileIndex[w];

		var m = !UCPath.match(chr1);
		UCPath = UCPath.replace(chr0,'');
		if(m) UCPath = UCPath.replace(chr1,'!');
		for(w = 0; w < UCFullPaths.length; ++w) {
			if((m ? UCFullPaths[w].replace(chr1,'!') : UCFullPaths[w]).replace(chr0,'') == UCPath) return cfb.FileIndex[w];
			if((m ? UCPaths[w].replace(chr1,'!') : UCPaths[w]).replace(chr0,'') == UCPath) return cfb.FileIndex[w];
		}
		return null;
	}
	/** CFB Constants */
	var MSSZ = 64; /* Mini Sector Size = 1<<6 */
	//var MSCSZ = 4096; /* Mini Stream Cutoff Size */
	/* 2.1 Compound File Sector Numbers and Types */
	var ENDOFCHAIN = -2;
	/* 2.2 Compound File Header */
	var HEADER_SIGNATURE = 'd0cf11e0a1b11ae1';
	var HEADER_SIG = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
	var HEADER_CLSID = '00000000000000000000000000000000';
	var consts = {
		/* 2.1 Compund File Sector Numbers and Types */
		MAXREGSECT: -6,
		DIFSECT: -4,
		FATSECT: -3,
		ENDOFCHAIN: ENDOFCHAIN,
		FREESECT: -1,
		/* 2.2 Compound File Header */
		HEADER_SIGNATURE: HEADER_SIGNATURE,
		HEADER_MINOR_VERSION: '3e00',
		MAXREGSID: -6,
		NOSTREAM: -1,
		HEADER_CLSID: HEADER_CLSID,
		/* 2.6.1 Compound File Directory Entry */
		EntryTypes: ['unknown','storage','stream','lockbytes','property','root']
	};

	function write_file(cfb/*:CFBContainer*/, filename/*:string*/, options/*:CFBWriteOpts*/)/*:void*/ {
		get_fs();
		var o = _write(cfb, options);
		/*:: if(typeof Buffer == 'undefined' || !Buffer.isBuffer(o) || !(o instanceof Buffer)) throw new Error("unreachable"); */
		fs.writeFileSync(filename, o);
	}

	function a2s(o/*:RawBytes*/)/*:string*/ {
		var out = new Array(o.length);
		for(var i = 0; i < o.length; ++i) out[i] = String.fromCharCode(o[i]);
		return out.join("");
	}

	function write(cfb/*:CFBContainer*/, options/*:CFBWriteOpts*/)/*:RawBytes|string*/ {
		var o = _write(cfb, options);
		switch(options && options.type || "buffer") {
			case "file": get_fs(); fs.writeFileSync(options.filename, (o/*:any*/)); return o;
			case "binary": return typeof o == "string" ? o : a2s(o);
			case "base64": return Base64_encode(typeof o == "string" ? o : a2s(o));
			case "buffer": if(has_buf) return Buffer.isBuffer(o) ? o : Buffer_from(o);
				/* falls through */
			case "array": return typeof o == "string" ? s2a(o) : o;
		}
		return o;
	}
	/* node < 8.1 zlib does not expose bytesRead, so default to pure JS */
	var _zlib;
	function use_zlib(zlib) { try {
		var InflateRaw = zlib.InflateRaw;
		var InflRaw = new InflateRaw();
		InflRaw._processChunk(new Uint8Array([3, 0]), InflRaw._finishFlushFlag);
		if(InflRaw.bytesRead) _zlib = zlib;
		else throw new Error("zlib does not expose bytesRead");
	} catch(e) {console.error("cannot use native zlib: " + (e.message || e)); } }

	function _inflateRawSync(payload, usz) {
		if(!_zlib) return _inflate(payload, usz);
		var InflateRaw = _zlib.InflateRaw;
		var InflRaw = new InflateRaw();
		var out = InflRaw._processChunk(payload.slice(payload.l), InflRaw._finishFlushFlag);
		payload.l += InflRaw.bytesRead;
		return out;
	}

	function _deflateRawSync(payload) {
		return _zlib ? _zlib.deflateRawSync(payload) : _deflate(payload);
	}
	var CLEN_ORDER = [ 16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15 ];

	/*  LEN_ID = [ 257, 258, 259, 260, 261, 262, 263, 264, 265, 266, 267, 268, 269, 270, 271, 272, 273, 274, 275, 276, 277, 278, 279, 280, 281, 282, 283, 284, 285 ]; */
	var LEN_LN = [   3,   4,   5,   6,   7,   8,   9,  10,  11,  13 , 15,  17,  19,  23,  27,  31,  35,  43,  51,  59,  67,  83,  99, 115, 131, 163, 195, 227, 258 ];

	/*  DST_ID = [  0,  1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13,  14,  15,  16,  17,  18,  19,   20,   21,   22,   23,   24,   25,   26,    27,    28,    29 ]; */
	var DST_LN = [  1,  2,  3,  4,  5,  7,  9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577 ];

	function bit_swap_8(n) { var t = (((((n<<1)|(n<<11)) & 0x22110) | (((n<<5)|(n<<15)) & 0x88440))); return ((t>>16) | (t>>8) |t)&0xFF; }

	var use_typed_arrays = typeof Uint8Array !== 'undefined';

	var bitswap8 = use_typed_arrays ? new Uint8Array(1<<8) : [];
	for(var q = 0; q < (1<<8); ++q) bitswap8[q] = bit_swap_8(q);

	function bit_swap_n(n, b) {
		var rev = bitswap8[n & 0xFF];
		if(b <= 8) return rev >>> (8-b);
		rev = (rev << 8) | bitswap8[(n>>8)&0xFF];
		if(b <= 16) return rev >>> (16-b);
		rev = (rev << 8) | bitswap8[(n>>16)&0xFF];
		return rev >>> (24-b);
	}

	/* helpers for unaligned bit reads */
	function read_bits_2(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 6 ? 0 : buf[h+1]<<8))>>>w)& 0x03; }
	function read_bits_3(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 5 ? 0 : buf[h+1]<<8))>>>w)& 0x07; }
	function read_bits_4(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 4 ? 0 : buf[h+1]<<8))>>>w)& 0x0F; }
	function read_bits_5(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 3 ? 0 : buf[h+1]<<8))>>>w)& 0x1F; }
	function read_bits_7(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 1 ? 0 : buf[h+1]<<8))>>>w)& 0x7F; }

	/* works up to n = 3 * 8 + 1 = 25 */
	function read_bits_n(buf, bl, n) {
		var w = (bl&7), h = (bl>>>3), f = ((1<<n)-1);
		var v = buf[h] >>> w;
		if(n < 8 - w) return v & f;
		v |= buf[h+1]<<(8-w);
		if(n < 16 - w) return v & f;
		v |= buf[h+2]<<(16-w);
		if(n < 24 - w) return v & f;
		v |= buf[h+3]<<(24-w);
		return v & f;
	}

	/* helpers for unaligned bit writes */
	function write_bits_3(buf, bl, v) { var w = bl & 7, h = bl >>> 3;
		if(w <= 5) buf[h] |= (v & 7) << w;
		else {
			buf[h] |= (v << w) & 0xFF;
			buf[h+1] = (v&7) >> (8-w);
		}
		return bl + 3;
	}

	function write_bits_1(buf, bl, v) {
		var w = bl & 7, h = bl >>> 3;
		v = (v&1) << w;
		buf[h] |= v;
		return bl + 1;
	}
	function write_bits_8(buf, bl, v) {
		var w = bl & 7, h = bl >>> 3;
		v <<= w;
		buf[h] |=  v & 0xFF; v >>>= 8;
		buf[h+1] = v;
		return bl + 8;
	}
	function write_bits_16(buf, bl, v) {
		var w = bl & 7, h = bl >>> 3;
		v <<= w;
		buf[h] |=  v & 0xFF; v >>>= 8;
		buf[h+1] = v & 0xFF;
		buf[h+2] = v >>> 8;
		return bl + 16;
	}

	/* until ArrayBuffer#realloc is a thing, fake a realloc */
	function realloc(b, sz/*:number*/) {
		var L = b.length, M = 2*L > sz ? 2*L : sz + 5, i = 0;
		if(L >= sz) return b;
		if(has_buf) {
			var o = new_unsafe_buf(M);
			// $FlowIgnore
			if(b.copy) b.copy(o);
			else for(; i < b.length; ++i) o[i] = b[i];
			return o;
		} else if(use_typed_arrays) {
			var a = new Uint8Array(M);
			if(a.set) a.set(b);
			else for(; i < L; ++i) a[i] = b[i];
			return a;
		}
		b.length = M;
		return b;
	}

	/* zero-filled arrays for older browsers */
	function zero_fill_array(n) {
		var o = new Array(n);
		for(var i = 0; i < n; ++i) o[i] = 0;
		return o;
	}

	/* build tree (used for literals and lengths) */
	function build_tree(clens, cmap, MAX/*:number*/)/*:number*/ {
		var maxlen = 1, w = 0, i = 0, j = 0, ccode = 0, L = clens.length;

		var bl_count  = use_typed_arrays ? new Uint16Array(32) : zero_fill_array(32);
		for(i = 0; i < 32; ++i) bl_count[i] = 0;

		for(i = L; i < MAX; ++i) clens[i] = 0;
		L = clens.length;

		var ctree = use_typed_arrays ? new Uint16Array(L) : zero_fill_array(L); // []

		/* build code tree */
		for(i = 0; i < L; ++i) {
			bl_count[(w = clens[i])]++;
			if(maxlen < w) maxlen = w;
			ctree[i] = 0;
		}
		bl_count[0] = 0;
		for(i = 1; i <= maxlen; ++i) bl_count[i+16] = (ccode = (ccode + bl_count[i-1])<<1);
		for(i = 0; i < L; ++i) {
			ccode = clens[i];
			if(ccode != 0) ctree[i] = bl_count[ccode+16]++;
		}

		/* cmap[maxlen + 4 bits] = (off&15) + (lit<<4) reverse mapping */
		var cleni = 0;
		for(i = 0; i < L; ++i) {
			cleni = clens[i];
			if(cleni != 0) {
				ccode = bit_swap_n(ctree[i], maxlen)>>(maxlen-cleni);
				for(j = (1<<(maxlen + 4 - cleni)) - 1; j>=0; --j)
					cmap[ccode|(j<<cleni)] = (cleni&15) | (i<<4);
			}
		}
		return maxlen;
	}

	/* Fixed Huffman */
	var fix_lmap = use_typed_arrays ? new Uint16Array(512) : zero_fill_array(512);
	var fix_dmap = use_typed_arrays ? new Uint16Array(32)  : zero_fill_array(32);
	if(!use_typed_arrays) {
		for(var i = 0; i < 512; ++i) fix_lmap[i] = 0;
		for(i = 0; i < 32; ++i) fix_dmap[i] = 0;
	}
	(function() {
		var dlens/*:Array<number>*/ = [];
		var i = 0;
		for(;i<32; i++) dlens.push(5);
		build_tree(dlens, fix_dmap, 32);

		var clens/*:Array<number>*/ = [];
		i = 0;
		for(; i<=143; i++) clens.push(8);
		for(; i<=255; i++) clens.push(9);
		for(; i<=279; i++) clens.push(7);
		for(; i<=287; i++) clens.push(8);
		build_tree(clens, fix_lmap, 288);
	})();var _deflateRaw = /*#__PURE__*/(function _deflateRawIIFE() {
		var DST_LN_RE = use_typed_arrays ? new Uint8Array(0x8000) : [];
		var j = 0, k = 0;
		for(; j < DST_LN.length - 1; ++j) {
			for(; k < DST_LN[j+1]; ++k) DST_LN_RE[k] = j;
		}
		for(;k < 32768; ++k) DST_LN_RE[k] = 29;

		var LEN_LN_RE = use_typed_arrays ? new Uint8Array(0x103) : [];
		for(j = 0, k = 0; j < LEN_LN.length - 1; ++j) {
			for(; k < LEN_LN[j+1]; ++k) LEN_LN_RE[k] = j;
		}

		function write_stored(data, out) {
			var boff = 0;
			while(boff < data.length) {
				var L = Math.min(0xFFFF, data.length - boff);
				var h = boff + L == data.length;
				out.write_shift(1, +h);
				out.write_shift(2, L);
				out.write_shift(2, (~L) & 0xFFFF);
				while(L-- > 0) out[out.l++] = data[boff++];
			}
			return out.l;
		}

		/* Fixed Huffman */
		function write_huff_fixed(data, out) {
			var bl = 0;
			var boff = 0;
			var addrs = use_typed_arrays ? new Uint16Array(0x8000) : [];
			while(boff < data.length) {
				var L = /* data.length - boff; */ Math.min(0xFFFF, data.length - boff);

				/* write a stored block for short data */
				if(L < 10) {
					bl = write_bits_3(out, bl, +!!(boff + L == data.length)); // jshint ignore:line
					if(bl & 7) bl += 8 - (bl & 7);
					out.l = (bl / 8) | 0;
					out.write_shift(2, L);
					out.write_shift(2, (~L) & 0xFFFF);
					while(L-- > 0) out[out.l++] = data[boff++];
					bl = out.l * 8;
					continue;
				}

				bl = write_bits_3(out, bl, +!!(boff + L == data.length) + 2); // jshint ignore:line
				var hash = 0;
				while(L-- > 0) {
					var d = data[boff];
					hash = ((hash << 5) ^ d) & 0x7FFF;

					var match = -1, mlen = 0;

					if((match = addrs[hash])) {
						match |= boff & ~0x7FFF;
						if(match > boff) match -= 0x8000;
						if(match < boff) while(data[match + mlen] == data[boff + mlen] && mlen < 250) ++mlen;
					}

					if(mlen > 2) {
						/* Copy Token  */
						d = LEN_LN_RE[mlen];
						if(d <= 22) bl = write_bits_8(out, bl, bitswap8[d+1]>>1) - 1;
						else {
							write_bits_8(out, bl, 3);
							bl += 5;
							write_bits_8(out, bl, bitswap8[d-23]>>5);
							bl += 3;
						}
						var len_eb = (d < 8) ? 0 : ((d - 4)>>2);
						if(len_eb > 0) {
							write_bits_16(out, bl, mlen - LEN_LN[d]);
							bl += len_eb;
						}

						d = DST_LN_RE[boff - match];
						bl = write_bits_8(out, bl, bitswap8[d]>>3);
						bl -= 3;

						var dst_eb = d < 4 ? 0 : (d-2)>>1;
						if(dst_eb > 0) {
							write_bits_16(out, bl, boff - match - DST_LN[d]);
							bl += dst_eb;
						}
						for(var q = 0; q < mlen; ++q) {
							addrs[hash] = boff & 0x7FFF;
							hash = ((hash << 5) ^ data[boff]) & 0x7FFF;
							++boff;
						}
						L-= mlen - 1;
					} else {
						/* Literal Token */
						if(d <= 143) d = d + 48;
						else bl = write_bits_1(out, bl, 1);
						bl = write_bits_8(out, bl, bitswap8[d]);
						addrs[hash] = boff & 0x7FFF;
						++boff;
					}
				}

				bl = write_bits_8(out, bl, 0) - 1;
			}
			out.l = ((bl + 7)/8)|0;
			return out.l;
		}
		return function _deflateRaw(data, out) {
			if(data.length < 8) return write_stored(data, out);
			return write_huff_fixed(data, out);
		};
	})();

	function _deflate(data) {
		var buf = new_buf(50+Math.floor(data.length*1.1));
		var off = _deflateRaw(data, buf);
		return buf.slice(0, off);
	}
	/* modified inflate function also moves original read head */

	var dyn_lmap = use_typed_arrays ? new Uint16Array(32768) : zero_fill_array(32768);
	var dyn_dmap = use_typed_arrays ? new Uint16Array(32768) : zero_fill_array(32768);
	var dyn_cmap = use_typed_arrays ? new Uint16Array(128)   : zero_fill_array(128);
	var dyn_len_1 = 1, dyn_len_2 = 1;

	/* 5.5.3 Expanding Huffman Codes */
	function dyn(data, boff/*:number*/) {
		/* nomenclature from RFC1951 refers to bit values; these are offset by the implicit constant */
		var _HLIT = read_bits_5(data, boff) + 257; boff += 5;
		var _HDIST = read_bits_5(data, boff) + 1; boff += 5;
		var _HCLEN = read_bits_4(data, boff) + 4; boff += 4;
		var w = 0;

		/* grab and store code lengths */
		var clens = use_typed_arrays ? new Uint8Array(19) : zero_fill_array(19);
		var ctree = [ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 ];
		var maxlen = 1;
		var bl_count =  use_typed_arrays ? new Uint8Array(8) : zero_fill_array(8);
		var next_code = use_typed_arrays ? new Uint8Array(8) : zero_fill_array(8);
		var L = clens.length; /* 19 */
		for(var i = 0; i < _HCLEN; ++i) {
			clens[CLEN_ORDER[i]] = w = read_bits_3(data, boff);
			if(maxlen < w) maxlen = w;
			bl_count[w]++;
			boff += 3;
		}

		/* build code tree */
		var ccode = 0;
		bl_count[0] = 0;
		for(i = 1; i <= maxlen; ++i) next_code[i] = ccode = (ccode + bl_count[i-1])<<1;
		for(i = 0; i < L; ++i) if((ccode = clens[i]) != 0) ctree[i] = next_code[ccode]++;
		/* cmap[7 bits from stream] = (off&7) + (lit<<3) */
		var cleni = 0;
		for(i = 0; i < L; ++i) {
			cleni = clens[i];
			if(cleni != 0) {
				ccode = bitswap8[ctree[i]]>>(8-cleni);
				for(var j = (1<<(7-cleni))-1; j>=0; --j) dyn_cmap[ccode|(j<<cleni)] = (cleni&7) | (i<<3);
			}
		}

		/* read literal and dist codes at once */
		var hcodes/*:Array<number>*/ = [];
		maxlen = 1;
		for(; hcodes.length < _HLIT + _HDIST;) {
			ccode = dyn_cmap[read_bits_7(data, boff)];
			boff += ccode & 7;
			switch((ccode >>>= 3)) {
				case 16:
					w = 3 + read_bits_2(data, boff); boff += 2;
					ccode = hcodes[hcodes.length - 1];
					while(w-- > 0) hcodes.push(ccode);
					break;
				case 17:
					w = 3 + read_bits_3(data, boff); boff += 3;
					while(w-- > 0) hcodes.push(0);
					break;
				case 18:
					w = 11 + read_bits_7(data, boff); boff += 7;
					while(w -- > 0) hcodes.push(0);
					break;
				default:
					hcodes.push(ccode);
					if(maxlen < ccode) maxlen = ccode;
					break;
			}
		}

		/* build literal / length trees */
		var h1 = hcodes.slice(0, _HLIT), h2 = hcodes.slice(_HLIT);
		for(i = _HLIT; i < 286; ++i) h1[i] = 0;
		for(i = _HDIST; i < 30; ++i) h2[i] = 0;
		dyn_len_1 = build_tree(h1, dyn_lmap, 286);
		dyn_len_2 = build_tree(h2, dyn_dmap, 30);
		return boff;
	}

	/* return [ data, bytesRead ] */
	function inflate(data, usz/*:number*/) {
		/* shortcircuit for empty buffer [0x03, 0x00] */
		if(data[0] == 3 && !(data[1] & 0x3)) { return [new_raw_buf(usz), 2]; }

		/* bit offset */
		var boff = 0;

		/* header includes final bit and type bits */
		var header = 0;

		var outbuf = new_unsafe_buf(usz ? usz : (1<<18));
		var woff = 0;
		var OL = outbuf.length>>>0;
		var max_len_1 = 0, max_len_2 = 0;

		while((header&1) == 0) {
			header = read_bits_3(data, boff); boff += 3;
			if((header >>> 1) == 0) {
				/* Stored block */
				if(boff & 7) boff += 8 - (boff&7);
				/* 2 bytes sz, 2 bytes bit inverse */
				var sz = data[boff>>>3] | data[(boff>>>3)+1]<<8;
				boff += 32;
				/* push sz bytes */
				if(sz > 0) {
					if(!usz && OL < woff + sz) { outbuf = realloc(outbuf, woff + sz); OL = outbuf.length; }
					while(sz-- > 0) { outbuf[woff++] = data[boff>>>3]; boff += 8; }
				}
				continue;
			} else if((header >> 1) == 1) {
				/* Fixed Huffman */
				max_len_1 = 9; max_len_2 = 5;
			} else {
				/* Dynamic Huffman */
				boff = dyn(data, boff);
				max_len_1 = dyn_len_1; max_len_2 = dyn_len_2;
			}
			for(;;) { // while(true) is apparently out of vogue in modern JS circles
				if(!usz && (OL < woff + 32767)) { outbuf = realloc(outbuf, woff + 32767); OL = outbuf.length; }
				/* ingest code and move read head */
				var bits = read_bits_n(data, boff, max_len_1);
				var code = (header>>>1) == 1 ? fix_lmap[bits] : dyn_lmap[bits];
				boff += code & 15; code >>>= 4;
				/* 0-255 are literals, 256 is end of block token, 257+ are copy tokens */
				if(((code>>>8)&0xFF) === 0) outbuf[woff++] = code;
				else if(code == 256) break;
				else {
					code -= 257;
					var len_eb = (code < 8) ? 0 : ((code-4)>>2); if(len_eb > 5) len_eb = 0;
					var tgt = woff + LEN_LN[code];
					/* length extra bits */
					if(len_eb > 0) {
						tgt += read_bits_n(data, boff, len_eb);
						boff += len_eb;
					}

					/* dist code */
					bits = read_bits_n(data, boff, max_len_2);
					code = (header>>>1) == 1 ? fix_dmap[bits] : dyn_dmap[bits];
					boff += code & 15; code >>>= 4;
					var dst_eb = (code < 4 ? 0 : (code-2)>>1);
					var dst = DST_LN[code];
					/* dist extra bits */
					if(dst_eb > 0) {
						dst += read_bits_n(data, boff, dst_eb);
						boff += dst_eb;
					}

					/* in the common case, manual byte copy is faster than TA set / Buffer copy */
					if(!usz && OL < tgt) { outbuf = realloc(outbuf, tgt + 100); OL = outbuf.length; }
					while(woff < tgt) { outbuf[woff] = outbuf[woff - dst]; ++woff; }
				}
			}
		}
		if(usz) return [outbuf, (boff+7)>>>3];
		return [outbuf.slice(0, woff), (boff+7)>>>3];
	}

	function _inflate(payload, usz) {
		var data = payload.slice(payload.l||0);
		var out = inflate(data, usz);
		payload.l += out[1];
		return out[0];
	}

	function warn_or_throw(wrn, msg) {
		if(wrn) { if(typeof console !== 'undefined') console.error(msg); }
		else throw new Error(msg);
	}

	function parse_zip(file/*:RawBytes*/, options/*:CFBReadOpts*/)/*:CFBContainer*/ {
		var blob/*:CFBlob*/ = /*::(*/file/*:: :any)*/;
		prep_blob(blob, 0);

		var FileIndex/*:CFBFileIndex*/ = [], FullPaths/*:Array<string>*/ = [];
		var o = {
			FileIndex: FileIndex,
			FullPaths: FullPaths
		};
		init_cfb(o, { root: options.root });

		/* find end of central directory, start just after signature */
		var i = blob.length - 4;
		while((blob[i] != 0x50 || blob[i+1] != 0x4b || blob[i+2] != 0x05 || blob[i+3] != 0x06) && i >= 0) --i;
		blob.l = i + 4;

		/* parse end of central directory */
		blob.l += 4;
		var fcnt = blob.read_shift(2);
		blob.l += 6;
		var start_cd = blob.read_shift(4);

		/* parse central directory */
		blob.l = start_cd;

		for(i = 0; i < fcnt; ++i) {
			/* trust local file header instead of CD entry */
			blob.l += 20;
			var csz = blob.read_shift(4);
			var usz = blob.read_shift(4);
			var namelen = blob.read_shift(2);
			var efsz = blob.read_shift(2);
			var fcsz = blob.read_shift(2);
			blob.l += 8;
			var offset = blob.read_shift(4);
			var EF = parse_extra_field(/*::(*/blob.slice(blob.l+namelen, blob.l+namelen+efsz)/*:: :any)*/);
			blob.l += namelen + efsz + fcsz;

			var L = blob.l;
			blob.l = offset + 4;
			parse_local_file(blob, csz, usz, o, EF);
			blob.l = L;
		}
		return o;
	}


	/* head starts just after local file header signature */
	function parse_local_file(blob/*:CFBlob*/, csz/*:number*/, usz/*:number*/, o/*:CFBContainer*/, EF) {
		/* [local file header] */
		blob.l += 2;
		var flags = blob.read_shift(2);
		var meth = blob.read_shift(2);
		var date = parse_dos_date(blob);

		if(flags & 0x2041) throw new Error("Unsupported ZIP encryption");
		var crc32 = blob.read_shift(4);
		var _csz = blob.read_shift(4);
		var _usz = blob.read_shift(4);

		var namelen = blob.read_shift(2);
		var efsz = blob.read_shift(2);

		// TODO: flags & (1<<11) // UTF8
		var name = ""; for(var i = 0; i < namelen; ++i) name += String.fromCharCode(blob[blob.l++]);
		if(efsz) {
			var ef = parse_extra_field(/*::(*/blob.slice(blob.l, blob.l + efsz)/*:: :any)*/);
			if((ef[0x5455]||{}).mt) date = ef[0x5455].mt;
			if(((EF||{})[0x5455]||{}).mt) date = EF[0x5455].mt;
		}
		blob.l += efsz;

		/* [encryption header] */

		/* [file data] */
		var data = blob.slice(blob.l, blob.l + _csz);
		switch(meth) {
			case 8: data = _inflateRawSync(blob, _usz); break;
			case 0: break; // TODO: scan for magic number
			default: throw new Error("Unsupported ZIP Compression method " + meth);
		}

		/* [data descriptor] */
		var wrn = false;
		if(flags & 8) {
			crc32 = blob.read_shift(4);
			if(crc32 == 0x08074b50) { crc32 = blob.read_shift(4); wrn = true; }
			_csz = blob.read_shift(4);
			_usz = blob.read_shift(4);
		}

		if(_csz != csz) warn_or_throw(wrn, "Bad compressed size: " + csz + " != " + _csz);
		if(_usz != usz) warn_or_throw(wrn, "Bad uncompressed size: " + usz + " != " + _usz);
		//var _crc32 = CRC32.buf(data, 0);
		//if((crc32>>0) != (_crc32>>0)) warn_or_throw(wrn, "Bad CRC32 checksum: " + crc32 + " != " + _crc32);
		cfb_add(o, name, data, {unsafe: true, mt: date});
	}
	function write_zip(cfb/*:CFBContainer*/, options/*:CFBWriteOpts*/)/*:RawBytes*/ {
		var _opts = options || {};
		var out = [], cdirs = [];
		var o/*:CFBlob*/ = new_buf(1);
		var method = (_opts.compression ? 8 : 0), flags = 0;
		var i = 0, j = 0;

		var start_cd = 0, fcnt = 0;
		var root = cfb.FullPaths[0], fp = root, fi = cfb.FileIndex[0];
		var crcs = [];
		var sz_cd = 0;

		for(i = 1; i < cfb.FullPaths.length; ++i) {
			fp = cfb.FullPaths[i].slice(root.length); fi = cfb.FileIndex[i];
			if(!fi.size || !fi.content || fp == "\u0001Sh33tJ5") continue;
			var start = start_cd;

			/* TODO: CP437 filename */
			var namebuf = new_buf(fp.length);
			for(j = 0; j < fp.length; ++j) namebuf.write_shift(1, fp.charCodeAt(j) & 0x7F);
			namebuf = namebuf.slice(0, namebuf.l);
			crcs[fcnt] = CRC32.buf(/*::((*/fi.content/*::||[]):any)*/, 0);

			var outbuf = fi.content/*::||[]*/;
			if(method == 8) outbuf = _deflateRawSync(outbuf);

			/* local file header */
			o = new_buf(30);
			o.write_shift(4, 0x04034b50);
			o.write_shift(2, 20);
			o.write_shift(2, flags);
			o.write_shift(2, method);
			/* TODO: last mod file time/date */
			if(fi.mt) write_dos_date(o, fi.mt);
			else o.write_shift(4, 0);
			o.write_shift(-4, crcs[fcnt]);
			o.write_shift(4,  outbuf.length);
			o.write_shift(4,  /*::(*/fi.content/*::||[])*/.length);
			o.write_shift(2, namebuf.length);
			o.write_shift(2, 0);

			start_cd += o.length;
			out.push(o);
			start_cd += namebuf.length;
			out.push(namebuf);

			/* TODO: extra fields? */

			/* TODO: encryption header ? */

			start_cd += outbuf.length;
			out.push(outbuf);

			/* central directory */
			o = new_buf(46);
			o.write_shift(4, 0x02014b50);
			o.write_shift(2, 0);
			o.write_shift(2, 20);
			o.write_shift(2, flags);
			o.write_shift(2, method);
			o.write_shift(4, 0); /* TODO: last mod file time/date */
			o.write_shift(-4, crcs[fcnt]);

			o.write_shift(4, outbuf.length);
			o.write_shift(4, /*::(*/fi.content/*::||[])*/.length);
			o.write_shift(2, namebuf.length);
			o.write_shift(2, 0);
			o.write_shift(2, 0);
			o.write_shift(2, 0);
			o.write_shift(2, 0);
			o.write_shift(4, 0);
			o.write_shift(4, start);

			sz_cd += o.l;
			cdirs.push(o);
			sz_cd += namebuf.length;
			cdirs.push(namebuf);
			++fcnt;
		}

		/* end of central directory */
		o = new_buf(22);
		o.write_shift(4, 0x06054b50);
		o.write_shift(2, 0);
		o.write_shift(2, 0);
		o.write_shift(2, fcnt);
		o.write_shift(2, fcnt);
		o.write_shift(4, sz_cd);
		o.write_shift(4, start_cd);
		o.write_shift(2, 0);

		return bconcat(([bconcat((out/*:any*/)), bconcat(cdirs), o]/*:any*/));
	}
	var ContentTypeMap = ({
		"htm": "text/html",
		"xml": "text/xml",

		"gif": "image/gif",
		"jpg": "image/jpeg",
		"png": "image/png",

		"mso": "application/x-mso",
		"thmx": "application/vnd.ms-officetheme",
		"sh33tj5": "application/octet-stream"
	}/*:any*/);

	function get_content_type(fi/*:CFBEntry*/, fp/*:string*/)/*:string*/ {
		if(fi.ctype) return fi.ctype;

		var ext = fi.name || "", m = ext.match(/\.([^\.]+)$/);
		if(m && ContentTypeMap[m[1]]) return ContentTypeMap[m[1]];

		if(fp) {
			m = (ext = fp).match(/[\.\\]([^\.\\])+$/);
			if(m && ContentTypeMap[m[1]]) return ContentTypeMap[m[1]];
		}

		return "application/octet-stream";
	}

	/* 76 character chunks TODO: intertwine encoding */
	function write_base64_76(bstr/*:string*/)/*:string*/ {
		var data = Base64_encode(bstr);
		var o = [];
		for(var i = 0; i < data.length; i+= 76) o.push(data.slice(i, i+76));
		return o.join("\r\n") + "\r\n";
	}

	/*
	Rules for QP:
		- escape =## applies for all non-display characters and literal "="
		- space or tab at end of line must be encoded
		- \r\n newlines can be preserved, but bare \r and \n must be escaped
		- lines must not exceed 76 characters, use soft breaks =\r\n

	TODO: Some files from word appear to write line extensions with bare equals:

	```
	<table class=3DMsoTableGrid border=3D1 cellspacing=3D0 cellpadding=3D0 width=
	="70%"
	```
	*/
	function write_quoted_printable(text/*:string*/)/*:string*/ {
		var encoded = text.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7E-\xFF=]/g, function(c) {
			var w = c.charCodeAt(0).toString(16).toUpperCase();
			return "=" + (w.length == 1 ? "0" + w : w);
		});

		encoded = encoded.replace(/ $/mg, "=20").replace(/\t$/mg, "=09");

		if(encoded.charAt(0) == "\n") encoded = "=0D" + encoded.slice(1);
		encoded = encoded.replace(/\r(?!\n)/mg, "=0D").replace(/\n\n/mg, "\n=0A").replace(/([^\r\n])\n/mg, "$1=0A");

		var o/*:Array<string>*/ = [], split = encoded.split("\r\n");
		for(var si = 0; si < split.length; ++si) {
			var str = split[si];
			if(str.length == 0) { o.push(""); continue; }
			for(var i = 0; i < str.length;) {
				var end = 76;
				var tmp = str.slice(i, i + end);
				if(tmp.charAt(end - 1) == "=") end --;
				else if(tmp.charAt(end - 2) == "=") end -= 2;
				else if(tmp.charAt(end - 3) == "=") end -= 3;
				tmp = str.slice(i, i + end);
				i += end;
				if(i < str.length) tmp += "=";
				o.push(tmp);
			}
		}

		return o.join("\r\n");
	}
	function parse_quoted_printable(data/*:Array<string>*/)/*:RawBytes*/ {
		var o = [];

		/* unify long lines */
		for(var di = 0; di < data.length; ++di) {
			var line = data[di];
			while(di <= data.length && line.charAt(line.length - 1) == "=") line = line.slice(0, line.length - 1) + data[++di];
			o.push(line);
		}

		/* decode */
		for(var oi = 0; oi < o.length; ++oi) o[oi] = o[oi].replace(/[=][0-9A-Fa-f]{2}/g, function($$) { return String.fromCharCode(parseInt($$.slice(1), 16)); });
		return s2a(o.join("\r\n"));
	}


	function parse_mime(cfb/*:CFBContainer*/, data/*:Array<string>*/, root/*:string*/)/*:void*/ {
		var fname = "", cte = "", ctype = "", fdata;
		var di = 0;
		for(;di < 10; ++di) {
			var line = data[di];
			if(!line || line.match(/^\s*$/)) break;
			var m = line.match(/^(.*?):\s*([^\s].*)$/);
			if(m) switch(m[1].toLowerCase()) {
				case "content-location": fname = m[2].trim(); break;
				case "content-type": ctype = m[2].trim(); break;
				case "content-transfer-encoding": cte = m[2].trim(); break;
			}
		}
		++di;
		switch(cte.toLowerCase()) {
			case 'base64': fdata = s2a(Base64_decode(data.slice(di).join(""))); break;
			case 'quoted-printable': fdata = parse_quoted_printable(data.slice(di)); break;
			default: throw new Error("Unsupported Content-Transfer-Encoding " + cte);
		}
		var file = cfb_add(cfb, fname.slice(root.length), fdata, {unsafe: true});
		if(ctype) file.ctype = ctype;
	}

	function parse_mad(file/*:RawBytes*/, options/*:CFBReadOpts*/)/*:CFBContainer*/ {
		if(a2s(file.slice(0,13)).toLowerCase() != "mime-version:") throw new Error("Unsupported MAD header");
		var root = (options && options.root || "");
		// $FlowIgnore
		var data = (has_buf && Buffer.isBuffer(file) ? file.toString("binary") : a2s(file)).split("\r\n");
		var di = 0, row = "";

		/* if root is not specified, scan for the common prefix */
		for(di = 0; di < data.length; ++di) {
			row = data[di];
			if(!/^Content-Location:/i.test(row)) continue;
			row = row.slice(row.indexOf("file"));
			if(!root) root = row.slice(0, row.lastIndexOf("/") + 1);
			if(row.slice(0, root.length) == root) continue;
			while(root.length > 0) {
				root = root.slice(0, root.length - 1);
				root = root.slice(0, root.lastIndexOf("/") + 1);
				if(row.slice(0,root.length) == root) break;
			}
		}

		var mboundary = (data[1] || "").match(/boundary="(.*?)"/);
		if(!mboundary) throw new Error("MAD cannot find boundary");
		var boundary = "--" + (mboundary[1] || "");

		var FileIndex/*:CFBFileIndex*/ = [], FullPaths/*:Array<string>*/ = [];
		var o = {
			FileIndex: FileIndex,
			FullPaths: FullPaths
		};
		init_cfb(o);
		var start_di, fcnt = 0;
		for(di = 0; di < data.length; ++di) {
			var line = data[di];
			if(line !== boundary && line !== boundary + "--") continue;
			if(fcnt++) parse_mime(o, data.slice(start_di, di), root);
			start_di = di;
		}
		return o;
	}

	function write_mad(cfb/*:CFBContainer*/, options/*:CFBWriteOpts*/)/*:string*/ {
		var opts = options || {};
		var boundary = opts.boundary || "SheetJS";
		boundary = '------=' + boundary;

		var out = [
			'MIME-Version: 1.0',
			'Content-Type: multipart/related; boundary="' + boundary.slice(2) + '"',
			'',
			'',
			''
		];

		var root = cfb.FullPaths[0], fp = root, fi = cfb.FileIndex[0];
		for(var i = 1; i < cfb.FullPaths.length; ++i) {
			fp = cfb.FullPaths[i].slice(root.length);
			fi = cfb.FileIndex[i];
			if(!fi.size || !fi.content || fp == "\u0001Sh33tJ5") continue;

			/* Normalize filename */
			fp = fp.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7E-\xFF]/g, function(c) {
				return "_x" + c.charCodeAt(0).toString(16) + "_";
			}).replace(/[\u0080-\uFFFF]/g, function(u) {
				return "_u" + u.charCodeAt(0).toString(16) + "_";
			});

			/* Extract content as binary string */
			var ca = fi.content;
			// $FlowIgnore
			var cstr = has_buf && Buffer.isBuffer(ca) ? ca.toString("binary") : a2s(ca);

			/* 4/5 of first 1024 chars ascii -> quoted printable, else base64 */
			var dispcnt = 0, L = Math.min(1024, cstr.length), cc = 0;
			for(var csl = 0; csl <= L; ++csl) if((cc=cstr.charCodeAt(csl)) >= 0x20 && cc < 0x80) ++dispcnt;
			var qp = dispcnt >= L * 4 / 5;

			out.push(boundary);
			out.push('Content-Location: ' + (opts.root || 'file:///C:/SheetJS/') + fp);
			out.push('Content-Transfer-Encoding: ' + (qp ? 'quoted-printable' : 'base64'));
			out.push('Content-Type: ' + get_content_type(fi, fp));
			out.push('');

			out.push(qp ? write_quoted_printable(cstr) : write_base64_76(cstr));
		}
		out.push(boundary + '--\r\n');
		return out.join("\r\n");
	}
	function cfb_new(opts/*:?any*/)/*:CFBContainer*/ {
		var o/*:CFBContainer*/ = ({}/*:any*/);
		init_cfb(o, opts);
		return o;
	}

	function cfb_add(cfb/*:CFBContainer*/, name/*:string*/, content/*:?RawBytes*/, opts/*:?any*/)/*:CFBEntry*/ {
		var unsafe = opts && opts.unsafe;
		if(!unsafe) init_cfb(cfb);
		var file = !unsafe && CFB.find(cfb, name);
		if(!file) {
			var fpath/*:string*/ = cfb.FullPaths[0];
			if(name.slice(0, fpath.length) == fpath) fpath = name;
			else {
				if(fpath.slice(-1) != "/") fpath += "/";
				fpath = (fpath + name).replace("//","/");
			}
			file = ({name: filename(name), type: 2}/*:any*/);
			cfb.FileIndex.push(file);
			cfb.FullPaths.push(fpath);
			if(!unsafe) CFB.utils.cfb_gc(cfb);
		}
		/*:: if(!file) throw new Error("unreachable"); */
		file.content = (content/*:any*/);
		file.size = content ? content.length : 0;
		if(opts) {
			if(opts.CLSID) file.clsid = opts.CLSID;
			if(opts.mt) file.mt = opts.mt;
			if(opts.ct) file.ct = opts.ct;
		}
		return file;
	}

	function cfb_del(cfb/*:CFBContainer*/, name/*:string*/)/*:boolean*/ {
		init_cfb(cfb);
		var file = CFB.find(cfb, name);
		if(file) for(var j = 0; j < cfb.FileIndex.length; ++j) if(cfb.FileIndex[j] == file) {
			cfb.FileIndex.splice(j, 1);
			cfb.FullPaths.splice(j, 1);
			return true;
		}
		return false;
	}

	function cfb_mov(cfb/*:CFBContainer*/, old_name/*:string*/, new_name/*:string*/)/*:boolean*/ {
		init_cfb(cfb);
		var file = CFB.find(cfb, old_name);
		if(file) for(var j = 0; j < cfb.FileIndex.length; ++j) if(cfb.FileIndex[j] == file) {
			cfb.FileIndex[j].name = filename(new_name);
			cfb.FullPaths[j] = new_name;
			return true;
		}
		return false;
	}

	function cfb_gc(cfb/*:CFBContainer*/)/*:void*/ { rebuild_cfb(cfb, true); }

	exports.find = find;
	exports.read = read;
	exports.parse = parse;
	exports.write = write;
	exports.writeFile = write_file;
	exports.utils = {
		cfb_new: cfb_new,
		cfb_add: cfb_add,
		cfb_del: cfb_del,
		cfb_mov: cfb_mov,
		cfb_gc: cfb_gc,
		ReadShift: ReadShift,
		CheckField: CheckField,
		prep_blob: prep_blob,
		bconcat: bconcat,
		use_zlib: use_zlib,
		_deflateRaw: _deflate,
		_inflateRaw: _inflate,
		consts: consts
	};

	return exports;
	})();

	/* normalize data for blob ctor */
	function blobify(data) {
		if(typeof data === "string") return s2ab(data);
		if(Array.isArray(data)) return a2u(data);
		return data;
	}
	/* write or download file */
	function write_dl(fname/*:string*/, payload/*:any*/, enc/*:?string*/) {
		if(typeof Deno !== 'undefined') {
			/* in this spot, it's safe to assume typed arrays and TextEncoder/TextDecoder exist */
			if(enc && typeof payload == "string") switch(enc) {
				case "utf8": payload = new TextEncoder(enc).encode(payload); break;
				case "binary": payload = s2ab(payload); break;
				/* TODO: binary equivalent */
				default: throw new Error("Unsupported encoding " + enc);
			}
			return Deno.writeFileSync(fname, payload);
		}
		var data = (enc == "utf8") ? utf8write(payload) : payload;
		/*:: declare var IE_SaveFile: any; */
		if(typeof IE_SaveFile !== 'undefined') return IE_SaveFile(data, fname);
		if(typeof Blob !== 'undefined') {
			var blob = new Blob([blobify(data)], {type:"application/octet-stream"});
			/*:: declare var navigator: any; */
			if(typeof navigator !== 'undefined' && navigator.msSaveBlob) return navigator.msSaveBlob(blob, fname);
			/*:: declare var saveAs: any; */
			if(typeof saveAs !== 'undefined') return saveAs(blob, fname);
			if(typeof URL !== 'undefined' && typeof document !== 'undefined' && document.createElement && URL.createObjectURL) {
				var url = URL.createObjectURL(blob);
				/*:: declare var chrome: any; */
				if(typeof chrome === 'object' && typeof (chrome.downloads||{}).download == "function") {
					if(URL.revokeObjectURL && typeof setTimeout !== 'undefined') setTimeout(function() { URL.revokeObjectURL(url); }, 60000);
					return chrome.downloads.download({ url: url, filename: fname, saveAs: true});
				}
				var a = document.createElement("a");
				if(a.download != null) {
					/*:: if(document.body == null) throw new Error("unreachable"); */
					a.download = fname; a.href = url; document.body.appendChild(a); a.click();
					/*:: if(document.body == null) throw new Error("unreachable"); */ document.body.removeChild(a);
					if(URL.revokeObjectURL && typeof setTimeout !== 'undefined') setTimeout(function() { URL.revokeObjectURL(url); }, 60000);
					return url;
				}
			}
		}
		// $FlowIgnore
		if(typeof $ !== 'undefined' && typeof File !== 'undefined' && typeof Folder !== 'undefined') try { // extendscript
			// $FlowIgnore
			var out = File(fname); out.open("w"); out.encoding = "binary";
			if(Array.isArray(payload)) payload = a2s(payload);
			out.write(payload); out.close(); return payload;
		} catch(e) { if(!e.message || !e.message.match(/onstruct/)) throw e; }
		throw new Error("cannot save file " + fname);
	}
	function keys(o/*:any*/)/*:Array<any>*/ {
		var ks = Object.keys(o), o2 = [];
		for(var i = 0; i < ks.length; ++i) if(Object.prototype.hasOwnProperty.call(o, ks[i])) o2.push(ks[i]);
		return o2;
	}

	function evert(obj/*:any*/)/*:EvertType*/ {
		var o = ([]/*:any*/), K = keys(obj);
		for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = K[i];
		return o;
	}

	function evert_num(obj/*:any*/)/*:EvertNumType*/ {
		var o = ([]/*:any*/), K = keys(obj);
		for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = parseInt(K[i],10);
		return o;
	}

	function evert_arr(obj/*:any*/)/*:EvertArrType*/ {
		var o/*:EvertArrType*/ = ([]/*:any*/), K = keys(obj);
		for(var i = 0; i !== K.length; ++i) {
			if(o[obj[K[i]]] == null) o[obj[K[i]]] = [];
			o[obj[K[i]]].push(K[i]);
		}
		return o;
	}

	var basedate = /*#__PURE__*/new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
	function datenum(v/*:Date*/, date1904/*:?boolean*/)/*:number*/ {
		var epoch = /*#__PURE__*/v.getTime();
		if(date1904) epoch -= 1462*24*60*60*1000;
		var dnthresh = /*#__PURE__*/basedate.getTime() + (/*#__PURE__*/v.getTimezoneOffset() - /*#__PURE__*/basedate.getTimezoneOffset()) * 60000;
		return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
	}

	var good_pd_date_1 = /*#__PURE__*/new Date('2017-02-19T19:06:09.000Z');
	var good_pd_date = /*#__PURE__*/isNaN(/*#__PURE__*/good_pd_date_1.getFullYear()) ? /*#__PURE__*/new Date('2/19/17') : good_pd_date_1;
	var good_pd = /*#__PURE__*/good_pd_date.getFullYear() == 2017;
	/* parses a date as a local date */
	function parseDate(str/*:string|Date*/, fixdate/*:?number*/)/*:Date*/ {
		var d = new Date(str);
		if(good_pd) {
			/*:: if(fixdate == null) fixdate = 0; */
			if(fixdate > 0) d.setTime(d.getTime() + d.getTimezoneOffset() * 60 * 1000);
			else if(fixdate < 0) d.setTime(d.getTime() - d.getTimezoneOffset() * 60 * 1000);
			return d;
		}
		if(str instanceof Date) return str;
		if(good_pd_date.getFullYear() == 1917 && !isNaN(d.getFullYear())) {
			var s = d.getFullYear();
			if(str.indexOf("" + s) > -1) return d;
			d.setFullYear(d.getFullYear() + 100); return d;
		}
		var n = str.match(/\d+/g)||["2017","2","19","0","0","0"];
		var out = new Date(+n[0], +n[1] - 1, +n[2], (+n[3]||0), (+n[4]||0), (+n[5]||0));
		if(str.indexOf("Z") > -1) out = new Date(out.getTime() - out.getTimezoneOffset() * 60 * 1000);
		return out;
	}

	function dup(o/*:any*/)/*:any*/ {
		if(typeof JSON != 'undefined' && !Array.isArray(o)) return JSON.parse(JSON.stringify(o));
		if(typeof o != 'object' || o == null) return o;
		if(o instanceof Date) return new Date(o.getTime());
		var out = {};
		for(var k in o) if(Object.prototype.hasOwnProperty.call(o, k)) out[k] = dup(o[k]);
		return out;
	}

	function fill(c/*:string*/,l/*:number*/)/*:string*/ { var o = ""; while(o.length < l) o+=c; return o; }

	/* TODO: stress test */
	function fuzzynum(s/*:string*/)/*:number*/ {
		var v/*:number*/ = Number(s);
		if(!isNaN(v)) return isFinite(v) ? v : NaN;
		if(!/\d/.test(s)) return v;
		var wt = 1;
		var ss = s.replace(/([\d]),([\d])/g,"$1$2").replace(/[$]/g,"").replace(/[%]/g, function() { wt *= 100; return "";});
		if(!isNaN(v = Number(ss))) return v / wt;
		ss = ss.replace(/[(](.*)[)]/,function($$, $1) { wt = -wt; return $1;});
		if(!isNaN(v = Number(ss))) return v / wt;
		return v;
	}
	var lower_months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
	function fuzzydate(s/*:string*/)/*:Date*/ {
		var o = new Date(s), n = new Date(NaN);
		var y = o.getYear(), m = o.getMonth(), d = o.getDate();
		if(isNaN(d)) return n;
		var lower = s.toLowerCase();
		if(lower.match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) {
			lower = lower.replace(/[^a-z]/g,"").replace(/([^a-z]|^)[ap]m?([^a-z]|$)/,"");
			if(lower.length > 3 && lower_months.indexOf(lower) == -1) return n;
		} else if(lower.match(/[a-z]/)) return n;
		if(y < 0 || y > 8099) return n;
		if((m > 0 || d > 1) && y != 101) return o;
		if(s.match(/[^-0-9:,\/\\]/)) return n;
		return o;
	}

	function zip_add_file(zip, path, content) {
		if(zip.FullPaths) {
			if(typeof content == "string") {
				var res;
				if(has_buf) res = Buffer_from(content);
				/* TODO: investigate performance in Edge 13 */
				//else if(typeof TextEncoder !== "undefined") res = new TextEncoder().encode(content);
				else res = utf8decode(content);
				return CFB.utils.cfb_add(zip, path, res);
			}
			CFB.utils.cfb_add(zip, path, content);
		}
		else zip.file(path, content);
	}

	function zip_new() { return CFB.utils.cfb_new(); }
	var XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

	var encodings = {
		'&quot;': '"',
		'&apos;': "'",
		'&gt;': '>',
		'&lt;': '<',
		'&amp;': '&'
	};
	var rencoding = /*#__PURE__*/evert(encodings);

	var decregex=/[&<>'"]/g, charegex = /[\u0000-\u0008\u000b-\u001f]/g;
	function escapexml(text/*:string*/)/*:string*/{
		var s = text + '';
		return s.replace(decregex, function(y) { return rencoding[y]; }).replace(charegex,function(s) { return "_x" + ("000"+s.charCodeAt(0).toString(16)).slice(-4) + "_";});
	}

	var htmlcharegex = /[\u0000-\u001f]/g;
	function escapehtml(text/*:string*/)/*:string*/{
		var s = text + '';
		return s.replace(decregex, function(y) { return rencoding[y]; }).replace(/\n/g, "<br/>").replace(htmlcharegex,function(s) { return "&#x" + ("000"+s.charCodeAt(0).toString(16)).slice(-4) + ";"; });
	}

	function utf8reada(orig/*:string*/)/*:string*/ {
		var out = "", i = 0, c = 0, d = 0, e = 0, f = 0, w = 0;
		while (i < orig.length) {
			c = orig.charCodeAt(i++);
			if (c < 128) { out += String.fromCharCode(c); continue; }
			d = orig.charCodeAt(i++);
			if (c>191 && c<224) { f = ((c & 31) << 6); f |= (d & 63); out += String.fromCharCode(f); continue; }
			e = orig.charCodeAt(i++);
			if (c < 240) { out += String.fromCharCode(((c & 15) << 12) | ((d & 63) << 6) | (e & 63)); continue; }
			f = orig.charCodeAt(i++);
			w = (((c & 7) << 18) | ((d & 63) << 12) | ((e & 63) << 6) | (f & 63))-65536;
			out += String.fromCharCode(0xD800 + ((w>>>10)&1023));
			out += String.fromCharCode(0xDC00 + (w&1023));
		}
		return out;
	}

	function utf8readb(data) {
		var out = new_raw_buf(2*data.length), w, i, j = 1, k = 0, ww=0, c;
		for(i = 0; i < data.length; i+=j) {
			j = 1;
			if((c=data.charCodeAt(i)) < 128) w = c;
			else if(c < 224) { w = (c&31)*64+(data.charCodeAt(i+1)&63); j=2; }
			else if(c < 240) { w=(c&15)*4096+(data.charCodeAt(i+1)&63)*64+(data.charCodeAt(i+2)&63); j=3; }
			else { j = 4;
				w = (c & 7)*262144+(data.charCodeAt(i+1)&63)*4096+(data.charCodeAt(i+2)&63)*64+(data.charCodeAt(i+3)&63);
				w -= 65536; ww = 0xD800 + ((w>>>10)&1023); w = 0xDC00 + (w&1023);
			}
			if(ww !== 0) { out[k++] = ww&255; out[k++] = ww>>>8; ww = 0; }
			out[k++] = w%256; out[k++] = w>>>8;
		}
		return out.slice(0,k).toString('ucs2');
	}

	function utf8readc(data) { return Buffer_from(data, 'binary').toString('utf8'); }

	var utf8corpus = "foo bar baz\u00e2\u0098\u0083\u00f0\u009f\u008d\u00a3";
	var utf8read = has_buf && (/*#__PURE__*/utf8readc(utf8corpus) == /*#__PURE__*/utf8reada(utf8corpus) && utf8readc || /*#__PURE__*/utf8readb(utf8corpus) == /*#__PURE__*/utf8reada(utf8corpus) && utf8readb) || utf8reada;

	var utf8write/*:StringConv*/ = has_buf ? function(data) { return Buffer_from(data, 'utf8').toString("binary"); } : function(orig/*:string*/)/*:string*/ {
		var out/*:Array<string>*/ = [], i = 0, c = 0, d = 0;
		while(i < orig.length) {
			c = orig.charCodeAt(i++);
			switch(true) {
				case c < 128: out.push(String.fromCharCode(c)); break;
				case c < 2048:
					out.push(String.fromCharCode(192 + (c >> 6)));
					out.push(String.fromCharCode(128 + (c & 63)));
					break;
				case c >= 55296 && c < 57344:
					c -= 55296; d = orig.charCodeAt(i++) - 56320 + (c<<10);
					out.push(String.fromCharCode(240 + ((d >>18) & 7)));
					out.push(String.fromCharCode(144 + ((d >>12) & 63)));
					out.push(String.fromCharCode(128 + ((d >> 6) & 63)));
					out.push(String.fromCharCode(128 + (d & 63)));
					break;
				default:
					out.push(String.fromCharCode(224 + (c >> 12)));
					out.push(String.fromCharCode(128 + ((c >> 6) & 63)));
					out.push(String.fromCharCode(128 + (c & 63)));
			}
		}
		return out.join("");
	};

	var htmldecode/*:{(s:string):string}*/ = /*#__PURE__*/(function() {
		var entities/*:Array<[RegExp, string]>*/ = [
			['nbsp', ' '], ['middot', '·'],
			['quot', '"'], ['apos', "'"], ['gt',   '>'], ['lt',   '<'], ['amp',  '&']
		].map(function(x/*:[string, string]*/) { return [new RegExp('&' + x[0] + ';', "ig"), x[1]]; });
		return function htmldecode(str/*:string*/)/*:string*/ {
			var o = str
					// Remove new lines and spaces from start of content
					.replace(/^[\t\n\r ]+/, "")
					// Remove new lines and spaces from end of content
					.replace(/[\t\n\r ]+$/,"")
					// Added line which removes any white space characters after and before html tags
					.replace(/>\s+/g,">").replace(/\s+</g,"<")
					// Replace remaining new lines and spaces with space
					.replace(/[\t\n\r ]+/g, " ")
					// Replace <br> tags with new lines
					.replace(/<\s*[bB][rR]\s*\/?>/g,"\n")
					// Strip HTML elements
					.replace(/<[^>]*>/g,"");
			for(var i = 0; i < entities.length; ++i) o = o.replace(entities[i][0], entities[i][1]);
			return o;
		};
	})();

	var wtregex = /(^\s|\s$|\n)/;
	function writetag(f/*:string*/,g/*:string*/)/*:string*/ { return '<' + f + (g.match(wtregex)?' xml:space="preserve"' : "") + '>' + g + '</' + f + '>'; }

	function wxt_helper(h)/*:string*/ { return keys(h).map(function(k) { return " " + k + '="' + h[k] + '"';}).join(""); }
	function writextag(f/*:string*/,g/*:?string*/,h) { return '<' + f + ((h != null) ? wxt_helper(h) : "") + ((g != null) ? (g.match(wtregex)?' xml:space="preserve"' : "") + '>' + g + '</' + f : "/") + '>';}

	function write_w3cdtf(d/*:Date*/, t/*:?boolean*/)/*:string*/ { try { return d.toISOString().replace(/\.\d*/,""); } catch(e) { if(t) throw e; } return ""; }

	function write_vt(s, xlsx/*:?boolean*/)/*:string*/ {
		switch(typeof s) {
			case 'string':
				var o = writextag('vt:lpwstr', escapexml(s));
				if(xlsx) o = o.replace(/&quot;/g, "_x0022_");
				return o;
			case 'number': return writextag((s|0)==s?'vt:i4':'vt:r8', escapexml(String(s)));
			case 'boolean': return writextag('vt:bool',s?'true':'false');
		}
		if(s instanceof Date) return writextag('vt:filetime', write_w3cdtf(s));
		throw new Error("Unable to serialize " + s);
	}
	//var xlmlregex = /<(\/?)([a-z0-9]*:|)(\w+)[^>]*>/mg;

	var XMLNS = ({
		CORE_PROPS: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
		CUST_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
		EXT_PROPS: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
		CT: 'http://schemas.openxmlformats.org/package/2006/content-types',
		RELS: 'http://schemas.openxmlformats.org/package/2006/relationships',
		TCMNT: 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments',
		'dc': 'http://purl.org/dc/elements/1.1/',
		'dcterms': 'http://purl.org/dc/terms/',
		'dcmitype': 'http://purl.org/dc/dcmitype/',
		'mx': 'http://schemas.microsoft.com/office/mac/excel/2008/main',
		'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
		'sjs': 'http://schemas.openxmlformats.org/package/2006/sheetjs/core-properties',
		'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
		'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
		'xsd': 'http://www.w3.org/2001/XMLSchema'
	}/*:any*/);

	var XMLNS_main = [
		'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
		'http://purl.oclc.org/ooxml/spreadsheetml/main',
		'http://schemas.microsoft.com/office/excel/2006/main',
		'http://schemas.microsoft.com/office/excel/2006/2'
	];

	var XLMLNS = ({
		'o':    'urn:schemas-microsoft-com:office:office',
		'x':    'urn:schemas-microsoft-com:office:excel',
		'ss':   'urn:schemas-microsoft-com:office:spreadsheet',
		'dt':   'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882',
		'mv':   'http://macVmlSchemaUri',
		'v':    'urn:schemas-microsoft-com:vml',
		'html': 'http://www.w3.org/TR/REC-html40'
	}/*:any*/);
	function read_double_le(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ {
		var s = 1 - 2 * (b[idx + 7] >>> 7);
		var e = ((b[idx + 7] & 0x7f) << 4) + ((b[idx + 6] >>> 4) & 0x0f);
		var m = (b[idx+6]&0x0f);
		for(var i = 5; i >= 0; --i) m = m * 256 + b[idx + i];
		if(e == 0x7ff) return m == 0 ? (s * Infinity) : NaN;
		if(e == 0) e = -1022;
		else { e -= 1023; m += Math.pow(2,52); }
		return s * Math.pow(2, e - 52) * m;
	}

	function write_double_le(b/*:RawBytes|CFBlob*/, v/*:number*/, idx/*:number*/) {
		var bs = ((((v < 0) || (1/v == -Infinity)) ? 1 : 0) << 7), e = 0, m = 0;
		var av = bs ? (-v) : v;
		if(!isFinite(av)) { e = 0x7ff; m = isNaN(v) ? 0x6969 : 0; }
		else if(av == 0) e = m = 0;
		else {
			e = Math.floor(Math.log(av) / Math.LN2);
			m = av * Math.pow(2, 52 - e);
			if((e <= -1023) && (!isFinite(m) || (m < Math.pow(2,52)))) { e = -1022; }
			else { m -= Math.pow(2,52); e+=1023; }
		}
		for(var i = 0; i <= 5; ++i, m/=256) b[idx + i] = m & 0xff;
		b[idx + 6] = ((e & 0x0f) << 4) | (m & 0xf);
		b[idx + 7] = (e >> 4) | bs;
	}

	var ___toBuffer = function(bufs/*:Array<Array<RawBytes> >*/)/*:RawBytes*/ { var x=[],w=10240; for(var i=0;i<bufs[0].length;++i) if(bufs[0][i]) for(var j=0,L=bufs[0][i].length;j<L;j+=w) x.push.apply(x, bufs[0][i].slice(j,j+w)); return x; };
	var __toBuffer = has_buf ? function(bufs) { return (bufs[0].length > 0 && Buffer.isBuffer(bufs[0][0])) ? Buffer.concat(bufs[0].map(function(x) { return Buffer.isBuffer(x) ? x : Buffer_from(x); })) : ___toBuffer(bufs);} : ___toBuffer;

	var ___utf16le = function(b/*:RawBytes|CFBlob*/,s/*:number*/,e/*:number*/)/*:string*/ { var ss/*:Array<string>*/=[]; for(var i=s; i<e; i+=2) ss.push(String.fromCharCode(__readUInt16LE(b,i))); return ss.join("").replace(chr0,''); };
	var __utf16le = has_buf ? function(b/*:RawBytes|CFBlob*/,s/*:number*/,e/*:number*/)/*:string*/ { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___utf16le(b,s,e); return b.toString('utf16le',s,e).replace(chr0,'')/*.replace(chr1,'!')*/; } : ___utf16le;

	var ___hexlify = function(b/*:RawBytes|CFBlob*/,s/*:number*/,l/*:number*/)/*:string*/ { var ss/*:Array<string>*/=[]; for(var i=s; i<s+l; ++i) ss.push(("0" + b[i].toString(16)).slice(-2)); return ss.join(""); };
	var __hexlify = has_buf ? function(b/*:RawBytes|CFBlob*/,s/*:number*/,l/*:number*/)/*:string*/ { return Buffer.isBuffer(b)/*:: && b instanceof Buffer*/ ? b.toString('hex',s,s+l) : ___hexlify(b,s,l); } : ___hexlify;

	var ___utf8 = function(b/*:RawBytes|CFBlob*/,s/*:number*/,e/*:number*/) { var ss=[]; for(var i=s; i<e; i++) ss.push(String.fromCharCode(__readUInt8(b,i))); return ss.join(""); };
	var __utf8 = has_buf ? function utf8_b(b/*:RawBytes|CFBlob*/, s/*:number*/, e/*:number*/) { return (Buffer.isBuffer(b)/*:: && (b instanceof Buffer)*/) ? b.toString('utf8',s,e) : ___utf8(b,s,e); } : ___utf8;

	var ___lpstr = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
	var __lpstr = ___lpstr;

	var ___cpstr = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
	var __cpstr = ___cpstr;

	var ___lpwstr = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = 2*__readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
	var __lpwstr = ___lpwstr;

	var ___lpp4 = function lpp4_(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? __utf16le(b, i+4,i+4+len) : "";};
	var __lpp4 = ___lpp4;

	var ___8lpp4 = function(b/*:RawBytes|CFBlob*/,i/*:number*/) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len) : "";};
	var __8lpp4 = ___8lpp4;

	var ___double = function(b/*:RawBytes|CFBlob*/, idx/*:number*/) { return read_double_le(b, idx);};
	var __double = ___double;

	if(has_buf/*:: && typeof Buffer !== 'undefined'*/) {
		__lpstr = function lpstr_b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___lpstr(b, i); var len = b.readUInt32LE(i); return len > 0 ? b.toString('utf8',i+4,i+4+len-1) : "";};
		__cpstr = function cpstr_b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___cpstr(b, i); var len = b.readUInt32LE(i); return len > 0 ? b.toString('utf8',i+4,i+4+len-1) : "";};
		__lpwstr = function lpwstr_b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___lpwstr(b, i); var len = 2*b.readUInt32LE(i); return b.toString('utf16le',i+4,i+4+len-1);};
		__lpp4 = function lpp4_b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___lpp4(b, i); var len = b.readUInt32LE(i); return b.toString('utf16le',i+4,i+4+len);};
		__8lpp4 = function lpp4_8b(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(!Buffer.isBuffer(b)/*:: || !(b instanceof Buffer)*/) return ___8lpp4(b, i); var len = b.readUInt32LE(i); return b.toString('utf8',i+4,i+4+len);};
		__double = function double_(b/*:RawBytes|CFBlob*/, i/*:number*/) { if(Buffer.isBuffer(b)/*::&& b instanceof Buffer*/) return b.readDoubleLE(i); return ___double(b,i); };
	}

	var __readUInt8 = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return b[idx]; };
	var __readUInt16LE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return (b[idx+1]*(1<<8))+b[idx]; };
	var __readInt16LE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { var u = (b[idx+1]*(1<<8))+b[idx]; return (u < 0x8000) ? u : ((0xffff - u + 1) * -1); };
	var __readUInt32LE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return b[idx+3]*(1<<24)+(b[idx+2]<<16)+(b[idx+1]<<8)+b[idx]; };
	var __readInt32LE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return (b[idx+3]<<24)|(b[idx+2]<<16)|(b[idx+1]<<8)|b[idx]; };
	var __readInt32BE = function(b/*:RawBytes|CFBlob*/, idx/*:number*/)/*:number*/ { return (b[idx]<<24)|(b[idx+1]<<16)|(b[idx+2]<<8)|b[idx+3]; };

	function ReadShift(size/*:number*/, t/*:?string*/)/*:number|string*/ {
		var o="", oI/*:: :number = 0*/, oR, oo=[], w, vv, i, loc;
		switch(t) {
			case 'dbcs':
				loc = this.l;
				if(has_buf && Buffer.isBuffer(this)) o = this.slice(this.l, this.l+2*size).toString("utf16le");
				else for(i = 0; i < size; ++i) { o+=String.fromCharCode(__readUInt16LE(this, loc)); loc+=2; }
				size *= 2;
				break;

			case 'utf8': o = __utf8(this, this.l, this.l + size); break;
			case 'utf16le': size *= 2; o = __utf16le(this, this.l, this.l + size); break;

			case 'wstr':
				return ReadShift.call(this, size, 'dbcs');

			/* [MS-OLEDS] 2.1.4 LengthPrefixedAnsiString */
			case 'lpstr-ansi': o = __lpstr(this, this.l); size = 4 + __readUInt32LE(this, this.l); break;
			case 'lpstr-cp': o = __cpstr(this, this.l); size = 4 + __readUInt32LE(this, this.l); break;
			/* [MS-OLEDS] 2.1.5 LengthPrefixedUnicodeString */
			case 'lpwstr': o = __lpwstr(this, this.l); size = 4 + 2 * __readUInt32LE(this, this.l); break;
			/* [MS-OFFCRYPTO] 2.1.2 Length-Prefixed Padded Unicode String (UNICODE-LP-P4) */
			case 'lpp4': size = 4 +  __readUInt32LE(this, this.l); o = __lpp4(this, this.l); if(size & 0x02) size += 2; break;
			/* [MS-OFFCRYPTO] 2.1.3 Length-Prefixed UTF-8 String (UTF-8-LP-P4) */
			case '8lpp4': size = 4 +  __readUInt32LE(this, this.l); o = __8lpp4(this, this.l); if(size & 0x03) size += 4 - (size & 0x03); break;

			case 'cstr': size = 0; o = "";
				while((w=__readUInt8(this, this.l + size++))!==0) oo.push(_getchar(w));
				o = oo.join(""); break;
			case '_wstr': size = 0; o = "";
				while((w=__readUInt16LE(this,this.l +size))!==0){oo.push(_getchar(w));size+=2;}
				size+=2; o = oo.join(""); break;

			/* sbcs and dbcs support continue records in the SST way TODO codepages */
			case 'dbcs-cont': o = ""; loc = this.l;
				for(i = 0; i < size; ++i) {
					if(this.lens && this.lens.indexOf(loc) !== -1) {
						w = __readUInt8(this, loc);
						this.l = loc + 1;
						vv = ReadShift.call(this, size-i, w ? 'dbcs-cont' : 'sbcs-cont');
						return oo.join("") + vv;
					}
					oo.push(_getchar(__readUInt16LE(this, loc)));
					loc+=2;
				} o = oo.join(""); size *= 2; break;

			case 'cpstr':
			/* falls through */
			case 'sbcs-cont': o = ""; loc = this.l;
				for(i = 0; i != size; ++i) {
					if(this.lens && this.lens.indexOf(loc) !== -1) {
						w = __readUInt8(this, loc);
						this.l = loc + 1;
						vv = ReadShift.call(this, size-i, w ? 'dbcs-cont' : 'sbcs-cont');
						return oo.join("") + vv;
					}
					oo.push(_getchar(__readUInt8(this, loc)));
					loc+=1;
				} o = oo.join(""); break;

			default:
		switch(size) {
			case 1: oI = __readUInt8(this, this.l); this.l++; return oI;
			case 2: oI = (t === 'i' ? __readInt16LE : __readUInt16LE)(this, this.l); this.l += 2; return oI;
			case 4: case -4:
				if(t === 'i' || ((this[this.l+3] & 0x80)===0)) { oI = ((size > 0) ? __readInt32LE : __readInt32BE)(this, this.l); this.l += 4; return oI; }
				else { oR = __readUInt32LE(this, this.l); this.l += 4; } return oR;
			case 8: case -8:
				if(t === 'f') {
					if(size == 8) oR = __double(this, this.l);
					else oR = __double([this[this.l+7],this[this.l+6],this[this.l+5],this[this.l+4],this[this.l+3],this[this.l+2],this[this.l+1],this[this.l+0]], 0);
					this.l += 8; return oR;
				} else size = 8;
			/* falls through */
			case 16: o = __hexlify(this, this.l, size); break;
		}}
		this.l+=size; return o;
	}

	var __writeUInt32LE = function(b/*:RawBytes|CFBlob*/, val/*:number*/, idx/*:number*/)/*:void*/ { b[idx] = (val & 0xFF); b[idx+1] = ((val >>> 8) & 0xFF); b[idx+2] = ((val >>> 16) & 0xFF); b[idx+3] = ((val >>> 24) & 0xFF); };
	var __writeInt32LE  = function(b/*:RawBytes|CFBlob*/, val/*:number*/, idx/*:number*/)/*:void*/ { b[idx] = (val & 0xFF); b[idx+1] = ((val >> 8) & 0xFF); b[idx+2] = ((val >> 16) & 0xFF); b[idx+3] = ((val >> 24) & 0xFF); };
	var __writeUInt16LE = function(b/*:RawBytes|CFBlob*/, val/*:number*/, idx/*:number*/)/*:void*/ { b[idx] = (val & 0xFF); b[idx+1] = ((val >>> 8) & 0xFF); };

	function WriteShift(t/*:number*/, val/*:string|number*/, f/*:?string*/)/*:any*/ {
		var size = 0, i = 0;
		if(f === 'dbcs') {
			/*:: if(typeof val !== 'string') throw new Error("unreachable"); */
			for(i = 0; i != val.length; ++i) __writeUInt16LE(this, val.charCodeAt(i), this.l + 2 * i);
			size = 2 * val.length;
		} else if(f === 'sbcs') {
			{
				/*:: if(typeof val !== 'string') throw new Error("unreachable"); */
				val = val.replace(/[^\x00-\x7F]/g, "_");
				/*:: if(typeof val !== 'string') throw new Error("unreachable"); */
				for(i = 0; i != val.length; ++i) this[this.l + i] = (val.charCodeAt(i) & 0xFF);
			}
			size = val.length;
		} else if(f === 'hex') {
			for(; i < t; ++i) {
				/*:: if(typeof val !== "string") throw new Error("unreachable"); */
				this[this.l++] = (parseInt(val.slice(2*i, 2*i+2), 16)||0);
			} return this;
		} else if(f === 'utf16le') {
				/*:: if(typeof val !== "string") throw new Error("unreachable"); */
				var end/*:number*/ = Math.min(this.l + t, this.length);
				for(i = 0; i < Math.min(val.length, t); ++i) {
					var cc = val.charCodeAt(i);
					this[this.l++] = (cc & 0xff);
					this[this.l++] = (cc >> 8);
				}
				while(this.l < end) this[this.l++] = 0;
				return this;
		} else /*:: if(typeof val === 'number') */ switch(t) {
			case  1: size = 1; this[this.l] = val&0xFF; break;
			case  2: size = 2; this[this.l] = val&0xFF; val >>>= 8; this[this.l+1] = val&0xFF; break;
			case  3: size = 3; this[this.l] = val&0xFF; val >>>= 8; this[this.l+1] = val&0xFF; val >>>= 8; this[this.l+2] = val&0xFF; break;
			case  4: size = 4; __writeUInt32LE(this, val, this.l); break;
			case  8: size = 8; if(f === 'f') { write_double_le(this, val, this.l); break; }
			/* falls through */
			case 16: break;
			case -4: size = 4; __writeInt32LE(this, val, this.l); break;
		}
		this.l += size; return this;
	}

	function CheckField(hexstr/*:string*/, fld/*:string*/)/*:void*/ {
		var m = __hexlify(this,this.l,hexstr.length>>1);
		if(m !== hexstr) throw new Error(fld + 'Expected ' + hexstr + ' saw ' + m);
		this.l += hexstr.length>>1;
	}

	function prep_blob(blob, pos/*:number*/)/*:void*/ {
		blob.l = pos;
		blob.read_shift = /*::(*/ReadShift/*:: :any)*/;
		blob.chk = CheckField;
		blob.write_shift = WriteShift;
	}

	function new_buf(sz/*:number*/)/*:Block*/ {
		var o = new_raw_buf(sz);
		prep_blob(o, 0);
		return o;
	}
	function decode_row(rowstr/*:string*/)/*:number*/ { return parseInt(unfix_row(rowstr),10) - 1; }
	function encode_row(row/*:number*/)/*:string*/ { return "" + (row + 1); }
	function unfix_row(cstr/*:string*/)/*:string*/ { return cstr.replace(/\$(\d+)$/,"$1"); }

	function decode_col(colstr/*:string*/)/*:number*/ { var c = unfix_col(colstr), d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
	function encode_col(col/*:number*/)/*:string*/ { if(col < 0) throw new Error("invalid column " + col); var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; }
	function unfix_col(cstr/*:string*/)/*:string*/ { return cstr.replace(/^\$([A-Z])/,"$1"); }

	function split_cell(cstr/*:string*/)/*:Array<string>*/ { return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/,"$1,$2").split(","); }
	//function decode_cell(cstr/*:string*/)/*:CellAddress*/ { var splt = split_cell(cstr); return { c:decode_col(splt[0]), r:decode_row(splt[1]) }; }
	function decode_cell(cstr/*:string*/)/*:CellAddress*/ {
		var R = 0, C = 0;
		for(var i = 0; i < cstr.length; ++i) {
			var cc = cstr.charCodeAt(i);
			if(cc >= 48 && cc <= 57) R = 10 * R + (cc - 48);
			else if(cc >= 65 && cc <= 90) C = 26 * C + (cc - 64);
		}
		return { c: C - 1, r:R - 1 };
	}
	//function encode_cell(cell/*:CellAddress*/)/*:string*/ { return encode_col(cell.c) + encode_row(cell.r); }
	function encode_cell(cell/*:CellAddress*/)/*:string*/ {
		var col = cell.c + 1;
		var s="";
		for(; col; col=((col-1)/26)|0) s = String.fromCharCode(((col-1)%26) + 65) + s;
		return s + (cell.r + 1);
	}
	function decode_range(range/*:string*/)/*:Range*/ {
		var idx = range.indexOf(":");
		if(idx == -1) return { s: decode_cell(range), e: decode_cell(range) };
		return { s: decode_cell(range.slice(0, idx)), e: decode_cell(range.slice(idx + 1)) };
	}
	/*# if only one arg, it is assumed to be a Range.  If 2 args, both are cell addresses */
	function encode_range(cs/*:CellAddrSpec|Range*/,ce/*:?CellAddrSpec*/)/*:string*/ {
		if(typeof ce === 'undefined' || typeof ce === 'number') {
	/*:: if(!(cs instanceof Range)) throw "unreachable"; */
			return encode_range(cs.s, cs.e);
		}
	/*:: if((cs instanceof Range)) throw "unreachable"; */
		if(typeof cs !== 'string') cs = encode_cell((cs/*:any*/));
		if(typeof ce !== 'string') ce = encode_cell((ce/*:any*/));
	/*:: if(typeof cs !== 'string') throw "unreachable"; */
	/*:: if(typeof ce !== 'string') throw "unreachable"; */
		return cs == ce ? cs : cs + ":" + ce;
	}

	function safe_decode_range(range/*:string*/)/*:Range*/ {
		var o = {s:{c:0,r:0},e:{c:0,r:0}};
		var idx = 0, i = 0, cc = 0;
		var len = range.length;
		for(idx = 0; i < len; ++i) {
			if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
			idx = 26*idx + cc;
		}
		o.s.c = --idx;

		for(idx = 0; i < len; ++i) {
			if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
			idx = 10*idx + cc;
		}
		o.s.r = --idx;

		if(i === len || cc != 10) { o.e.c=o.s.c; o.e.r=o.s.r; return o; }
		++i;

		for(idx = 0; i != len; ++i) {
			if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
			idx = 26*idx + cc;
		}
		o.e.c = --idx;

		for(idx = 0; i != len; ++i) {
			if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
			idx = 10*idx + cc;
		}
		o.e.r = --idx;
		return o;
	}

	function safe_format_cell(cell/*:Cell*/, v/*:any*/) {
		var q = (cell.t == 'd' && v instanceof Date);
		if(cell.z != null) try { return (cell.w = SSF_format(cell.z, q ? datenum(v) : v)); } catch(e) { }
		try { return (cell.w = SSF_format((cell.XF||{}).numFmtId||(q ? 14 : 0),  q ? datenum(v) : v)); } catch(e) { return ''+v; }
	}

	function format_cell(cell/*:Cell*/, v/*:any*/, o/*:any*/) {
		if(cell == null || cell.t == null || cell.t == 'z') return "";
		if(cell.w !== undefined) return cell.w;
		if(cell.t == 'd' && !cell.z && o && o.dateNF) cell.z = o.dateNF;
		if(cell.t == "e") return BErr[cell.v] || cell.v;
		if(v == undefined) return safe_format_cell(cell, cell.v);
		return safe_format_cell(cell, v);
	}

	function sheet_to_workbook(sheet/*:Worksheet*/, opts)/*:Workbook*/ {
		var n = opts && opts.sheet ? opts.sheet : "Sheet1";
		var sheets = {}; sheets[n] = sheet;
		return { SheetNames: [n], Sheets: sheets };
	}

	function sheet_add_aoa(_ws/*:?Worksheet*/, data/*:AOA*/, opts/*:?any*/)/*:Worksheet*/ {
		var o = opts || {};
		var dense = _ws ? Array.isArray(_ws) : o.dense;
		var ws/*:Worksheet*/ = _ws || (dense ? ([]/*:any*/) : ({}/*:any*/));
		var _R = 0, _C = 0;
		if(ws && o.origin != null) {
			if(typeof o.origin == 'number') _R = o.origin;
			else {
				var _origin/*:CellAddress*/ = typeof o.origin == "string" ? decode_cell(o.origin) : o.origin;
				_R = _origin.r; _C = _origin.c;
			}
			if(!ws["!ref"]) ws["!ref"] = "A1:A1";
		}
		var range/*:Range*/ = ({s: {c:10000000, r:10000000}, e: {c:0, r:0}}/*:any*/);
		if(ws['!ref']) {
			var _range = safe_decode_range(ws['!ref']);
			range.s.c = _range.s.c;
			range.s.r = _range.s.r;
			range.e.c = Math.max(range.e.c, _range.e.c);
			range.e.r = Math.max(range.e.r, _range.e.r);
			if(_R == -1) range.e.r = _R = _range.e.r + 1;
		}
		for(var R = 0; R != data.length; ++R) {
			if(!data[R]) continue;
			if(!Array.isArray(data[R])) throw new Error("aoa_to_sheet expects an array of arrays");
			for(var C = 0; C != data[R].length; ++C) {
				if(typeof data[R][C] === 'undefined') continue;
				var cell/*:Cell*/ = ({v: data[R][C] }/*:any*/);
				var __R = _R + R, __C = _C + C;
				if(range.s.r > __R) range.s.r = __R;
				if(range.s.c > __C) range.s.c = __C;
				if(range.e.r < __R) range.e.r = __R;
				if(range.e.c < __C) range.e.c = __C;
				if(data[R][C] && typeof data[R][C] === 'object' && !Array.isArray(data[R][C]) && !(data[R][C] instanceof Date)) cell = data[R][C];
				else {
					if(Array.isArray(cell.v)) { cell.f = data[R][C][1]; cell.v = cell.v[0]; }
					if(cell.v === null) {
						if(cell.f) cell.t = 'n';
						else if(o.nullError) { cell.t = 'e'; cell.v = 0; }
						else if(!o.sheetStubs) continue;
						else cell.t = 'z';
					}
					else if(typeof cell.v === 'number') cell.t = 'n';
					else if(typeof cell.v === 'boolean') cell.t = 'b';
					else if(cell.v instanceof Date) {
						cell.z = o.dateNF || table_fmt[14];
						if(o.cellDates) { cell.t = 'd'; cell.w = SSF_format(cell.z, datenum(cell.v)); }
						else { cell.t = 'n'; cell.v = datenum(cell.v); cell.w = SSF_format(cell.z, cell.v); }
					}
					else cell.t = 's';
				}
				if(dense) {
					if(!ws[__R]) ws[__R] = [];
					if(ws[__R][__C] && ws[__R][__C].z) cell.z = ws[__R][__C].z;
					ws[__R][__C] = cell;
				} else {
					var cell_ref = encode_cell(({c:__C,r:__R}/*:any*/));
					if(ws[cell_ref] && ws[cell_ref].z) cell.z = ws[cell_ref].z;
					ws[cell_ref] = cell;
				}
			}
		}
		if(range.s.c < 10000000) ws['!ref'] = encode_range(range);
		return ws;
	}
	function aoa_to_sheet(data/*:AOA*/, opts/*:?any*/)/*:Worksheet*/ { return sheet_add_aoa(null, data, opts); }

	/* [MS-XLSB] 2.5.97.2 */
	var BErr = {
		/*::[*/0x00/*::]*/: "#NULL!",
		/*::[*/0x07/*::]*/: "#DIV/0!",
		/*::[*/0x0F/*::]*/: "#VALUE!",
		/*::[*/0x17/*::]*/: "#REF!",
		/*::[*/0x1D/*::]*/: "#NAME?",
		/*::[*/0x24/*::]*/: "#NUM!",
		/*::[*/0x2A/*::]*/: "#N/A",
		/*::[*/0x2B/*::]*/: "#GETTING_DATA",
		/*::[*/0xFF/*::]*/: "#WTF?"
	};

	/* Parts enumerated in OPC spec, MS-XLSB and MS-XLSX */
	/* 12.3 Part Summary <SpreadsheetML> */
	/* 14.2 Part Summary <DrawingML> */
	/* [MS-XLSX] 2.1 Part Enumerations ; [MS-XLSB] 2.1.7 Part Enumeration */
	var ct2type/*{[string]:string}*/ = ({
		/* Workbook */
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks",
		"application/vnd.ms-excel.sheet.macroEnabled.main+xml": "workbooks",
		"application/vnd.ms-excel.sheet.binary.macroEnabled.main": "workbooks",
		"application/vnd.ms-excel.addin.macroEnabled.main+xml": "workbooks",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "workbooks",

		/* Worksheet */
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml": "sheets",
		"application/vnd.ms-excel.worksheet": "sheets",
		"application/vnd.ms-excel.binIndexWs": "TODO", /* Binary Index */

		/* Chartsheet */
		"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml": "charts",
		"application/vnd.ms-excel.chartsheet": "charts",

		/* Macrosheet */
		"application/vnd.ms-excel.macrosheet+xml": "macros",
		"application/vnd.ms-excel.macrosheet": "macros",
		"application/vnd.ms-excel.intlmacrosheet": "TODO",
		"application/vnd.ms-excel.binIndexMs": "TODO", /* Binary Index */

		/* Dialogsheet */
		"application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml": "dialogs",
		"application/vnd.ms-excel.dialogsheet": "dialogs",

		/* Shared Strings */
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "strs",
		"application/vnd.ms-excel.sharedStrings": "strs",

		/* Styles */
		"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml": "styles",
		"application/vnd.ms-excel.styles": "styles",

		/* File Properties */
		"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
		"application/vnd.openxmlformats-officedocument.custom-properties+xml": "custprops",
		"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",

		/* Custom Data Properties */
		"application/vnd.openxmlformats-officedocument.customXmlProperties+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty": "TODO",

		/* Comments */
		"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "comments",
		"application/vnd.ms-excel.comments": "comments",
		"application/vnd.ms-excel.threadedcomments+xml": "threadedcomments",
		"application/vnd.ms-excel.person+xml": "people",

		/* Metadata (Stock/Geography and Dynamic Array) */
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml": "metadata",
		"application/vnd.ms-excel.sheetMetadata": "metadata",

		/* PivotTable */
		"application/vnd.ms-excel.pivotTable": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml": "TODO",

		/* Chart Objects */
		"application/vnd.openxmlformats-officedocument.drawingml.chart+xml": "TODO",

		/* Chart Colors */
		"application/vnd.ms-office.chartcolorstyle+xml": "TODO",

		/* Chart Style */
		"application/vnd.ms-office.chartstyle+xml": "TODO",

		/* Chart Advanced */
		"application/vnd.ms-office.chartex+xml": "TODO",

		/* Calculation Chain */
		"application/vnd.ms-excel.calcChain": "calcchains",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains",

		/* Printer Settings */
		"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings": "TODO",

		/* ActiveX */
		"application/vnd.ms-office.activeX": "TODO",
		"application/vnd.ms-office.activeX+xml": "TODO",

		/* Custom Toolbars */
		"application/vnd.ms-excel.attachedToolbars": "TODO",

		/* External Data Connections */
		"application/vnd.ms-excel.connections": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml": "TODO",

		/* External Links */
		"application/vnd.ms-excel.externalLink": "links",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml": "links",

		/* PivotCache */
		"application/vnd.ms-excel.pivotCacheDefinition": "TODO",
		"application/vnd.ms-excel.pivotCacheRecords": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml": "TODO",

		/* Query Table */
		"application/vnd.ms-excel.queryTable": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml": "TODO",

		/* Shared Workbook */
		"application/vnd.ms-excel.userNames": "TODO",
		"application/vnd.ms-excel.revisionHeaders": "TODO",
		"application/vnd.ms-excel.revisionLog": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml": "TODO",

		/* Single Cell Table */
		"application/vnd.ms-excel.tableSingleCells": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml": "TODO",

		/* Slicer */
		"application/vnd.ms-excel.slicer": "TODO",
		"application/vnd.ms-excel.slicerCache": "TODO",
		"application/vnd.ms-excel.slicer+xml": "TODO",
		"application/vnd.ms-excel.slicerCache+xml": "TODO",

		/* Sort Map */
		"application/vnd.ms-excel.wsSortMap": "TODO",

		/* Table */
		"application/vnd.ms-excel.table": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml": "TODO",

		/* Themes */
		"application/vnd.openxmlformats-officedocument.theme+xml": "themes",

		/* Theme Override */
		"application/vnd.openxmlformats-officedocument.themeOverride+xml": "TODO",

		/* Timeline */
		"application/vnd.ms-excel.Timeline+xml": "TODO", /* verify */
		"application/vnd.ms-excel.TimelineCache+xml": "TODO", /* verify */

		/* VBA */
		"application/vnd.ms-office.vbaProject": "vba",
		"application/vnd.ms-office.vbaProjectSignature": "TODO",

		/* Volatile Dependencies */
		"application/vnd.ms-office.volatileDependencies": "TODO",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml": "TODO",

		/* Control Properties */
		"application/vnd.ms-excel.controlproperties+xml": "TODO",

		/* Data Model */
		"application/vnd.openxmlformats-officedocument.model+data": "TODO",

		/* Survey */
		"application/vnd.ms-excel.Survey+xml": "TODO",

		/* Drawing */
		"application/vnd.openxmlformats-officedocument.drawing+xml": "drawings",
		"application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml": "TODO",
		"application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml": "TODO",

		/* VML */
		"application/vnd.openxmlformats-officedocument.vmlDrawing": "TODO",

		"application/vnd.openxmlformats-package.relationships+xml": "rels",
		"application/vnd.openxmlformats-officedocument.oleObject": "TODO",

		/* Image */
		"image/png": "TODO",

		"sheet": "js"
	}/*:any*/);

	var CT_LIST = {
			workbooks: {
				xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
				xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
				xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
				xlam: "application/vnd.ms-excel.addin.macroEnabled.main+xml",
				xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
			},
			strs: { /* Shared Strings */
				xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
				xlsb: "application/vnd.ms-excel.sharedStrings"
			},
			comments: { /* Comments */
				xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
				xlsb: "application/vnd.ms-excel.comments"
			},
			sheets: { /* Worksheet */
				xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
				xlsb: "application/vnd.ms-excel.worksheet"
			},
			charts: { /* Chartsheet */
				xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
				xlsb: "application/vnd.ms-excel.chartsheet"
			},
			dialogs: { /* Dialogsheet */
				xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
				xlsb: "application/vnd.ms-excel.dialogsheet"
			},
			macros: { /* Macrosheet (Excel 4.0 Macros) */
				xlsx: "application/vnd.ms-excel.macrosheet+xml",
				xlsb: "application/vnd.ms-excel.macrosheet"
			},
			metadata: { /* Metadata (Stock/Geography and Dynamic Array) */
				xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml",
				xlsb: "application/vnd.ms-excel.sheetMetadata"
			},
			styles: { /* Styles */
				xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
				xlsb: "application/vnd.ms-excel.styles"
			}
	};

	function new_ct()/*:any*/ {
		return ({
			workbooks:[], sheets:[], charts:[], dialogs:[], macros:[],
			rels:[], strs:[], comments:[], threadedcomments:[], links:[],
			coreprops:[], extprops:[], custprops:[], themes:[], styles:[],
			calcchains:[], vba: [], drawings: [], metadata: [], people:[],
			TODO:[], xmlns: "" }/*:any*/);
	}

	function write_ct(ct, opts)/*:string*/ {
		var type2ct/*{[string]:Array<string>}*/ = evert_arr(ct2type);

		var o/*:Array<string>*/ = [], v;
		o[o.length] = (XML_HEADER);
		o[o.length] = writextag('Types', null, {
			'xmlns': XMLNS.CT,
			'xmlns:xsd': XMLNS.xsd,
			'xmlns:xsi': XMLNS.xsi
		});

		o = o.concat([
			['xml', 'application/xml'],
			['bin', 'application/vnd.ms-excel.sheet.binary.macroEnabled.main'],
			['vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing'],
			['data', 'application/vnd.openxmlformats-officedocument.model+data'],
			/* from test files */
			['bmp', 'image/bmp'],
			['png', 'image/png'],
			['gif', 'image/gif'],
			['emf', 'image/x-emf'],
			['wmf', 'image/x-wmf'],
			['jpg', 'image/jpeg'], ['jpeg', 'image/jpeg'],
			['tif', 'image/tiff'], ['tiff', 'image/tiff'],
			['pdf', 'application/pdf'],
			['rels', 'application/vnd.openxmlformats-package.relationships+xml']
		].map(function(x) {
			return writextag('Default', null, {'Extension':x[0], 'ContentType': x[1]});
		}));

		/* only write first instance */
		var f1 = function(w) {
			if(ct[w] && ct[w].length > 0) {
				v = ct[w][0];
				o[o.length] = (writextag('Override', null, {
					'PartName': (v[0] == '/' ? "":"/") + v,
					'ContentType': CT_LIST[w][opts.bookType] || CT_LIST[w]['xlsx']
				}));
			}
		};

		/* book type-specific */
		var f2 = function(w) {
			(ct[w]||[]).forEach(function(v) {
				o[o.length] = (writextag('Override', null, {
					'PartName': (v[0] == '/' ? "":"/") + v,
					'ContentType': CT_LIST[w][opts.bookType] || CT_LIST[w]['xlsx']
				}));
			});
		};

		/* standard type */
		var f3 = function(t) {
			(ct[t]||[]).forEach(function(v) {
				o[o.length] = (writextag('Override', null, {
					'PartName': (v[0] == '/' ? "":"/") + v,
					'ContentType': type2ct[t][0]
				}));
			});
		};

		f1('workbooks');
		f2('sheets');
		f2('charts');
		f3('themes');
		['strs', 'styles'].forEach(f1);
		['coreprops', 'extprops', 'custprops'].forEach(f3);
		f3('vba');
		f3('comments');
		f3('threadedcomments');
		f3('drawings');
		f2('metadata');
		f3('people');
		if(o.length>2){ o[o.length] = ('</Types>'); o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}
	/* 9.3 Relationships */
	var RELS = ({
		WB: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		SHEET: "http://sheetjs.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
		HLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
		VML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
		XPATH: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
		XMISS: "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing",
		XLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
		CXML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
		CXMLP: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps",
		CMNT: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
		CORE_PROPS: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
		EXT_PROPS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
		CUST_PROPS: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
		SST: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
		STY: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
		THEME: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
		CHART: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
		CHARTEX: "http://schemas.microsoft.com/office/2014/relationships/chartEx",
		CS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
		WS: [
			"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
			"http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
		],
		DS: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
		MS: "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
		IMG: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		DRAW: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
		XLMETA: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
		TCMNT: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
		PEOPLE: "http://schemas.microsoft.com/office/2017/10/relationships/person",
		VBA: "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
	}/*:any*/);


	/* 9.3.3 Representing Relationships */
	function get_rels_path(file/*:string*/)/*:string*/ {
		var n = file.lastIndexOf("/");
		return file.slice(0,n+1) + '_rels/' + file.slice(n+1) + ".rels";
	}


	/* TODO */
	function write_rels(rels)/*:string*/ {
		var o = [XML_HEADER, writextag('Relationships', null, {
			//'xmlns:ns0': XMLNS.RELS,
			'xmlns': XMLNS.RELS
		})];
		keys(rels['!id']).forEach(function(rid) {
			o[o.length] = (writextag('Relationship', null, rels['!id'][rid]));
		});
		if(o.length>2){ o[o.length] = ('</Relationships>'); o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}

	function add_rels(rels, rId/*:number*/, f, type, relobj, targetmode/*:?string*/)/*:number*/ {
		if(!relobj) relobj = {};
		if(!rels['!id']) rels['!id'] = {};
		if(!rels['!idx']) rels['!idx'] = 1;
		if(rId < 0) for(rId = rels['!idx']; rels['!id']['rId' + rId]; ++rId){/* empty */}
		rels['!idx'] = rId + 1;
		relobj.Id = 'rId' + rId;
		relobj.Type = type;
		relobj.Target = f;
		if(targetmode) relobj.TargetMode = targetmode;
		else if([RELS.HLINK, RELS.XPATH, RELS.XMISS].indexOf(relobj.Type) > -1) relobj.TargetMode = "External";
		if(rels['!id'][relobj.Id]) throw new Error("Cannot rewrite rId " + rId);
		rels['!id'][relobj.Id] = relobj;
		rels[('/' + relobj.Target).replace("//","/")] = relobj;
		return rId;
	}

	/* ECMA-376 Part II 11.1 Core Properties Part */
	/* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
	var CORE_PROPS/*:Array<Array<string> >*/ = [
		["cp:category", "Category"],
		["cp:contentStatus", "ContentStatus"],
		["cp:keywords", "Keywords"],
		["cp:lastModifiedBy", "LastAuthor"],
		["cp:lastPrinted", "LastPrinted"],
		["cp:revision", "RevNumber"],
		["cp:version", "Version"],
		["dc:creator", "Author"],
		["dc:description", "Comments"],
		["dc:identifier", "Identifier"],
		["dc:language", "Language"],
		["dc:subject", "Subject"],
		["dc:title", "Title"],
		["dcterms:created", "CreatedDate", 'date'],
		["dcterms:modified", "ModifiedDate", 'date']
	];

	function cp_doit(f, g, h, o, p) {
		if(p[f] != null || g == null || g === "") return;
		p[f] = g;
		g = escapexml(g);
		o[o.length] = (h ? writextag(f,g,h) : writetag(f,g));
	}

	function write_core_props(cp, _opts) {
		var opts = _opts || {};
		var o = [XML_HEADER, writextag('cp:coreProperties', null, {
			//'xmlns': XMLNS.CORE_PROPS,
			'xmlns:cp': XMLNS.CORE_PROPS,
			'xmlns:dc': XMLNS.dc,
			'xmlns:dcterms': XMLNS.dcterms,
			'xmlns:dcmitype': XMLNS.dcmitype,
			'xmlns:xsi': XMLNS.xsi
		})], p = {};
		if(!cp && !opts.Props) return o.join("");

		if(cp) {
			if(cp.CreatedDate != null) cp_doit("dcterms:created", typeof cp.CreatedDate === "string" ? cp.CreatedDate : write_w3cdtf(cp.CreatedDate, opts.WTF), {"xsi:type":"dcterms:W3CDTF"}, o, p);
			if(cp.ModifiedDate != null) cp_doit("dcterms:modified", typeof cp.ModifiedDate === "string" ? cp.ModifiedDate : write_w3cdtf(cp.ModifiedDate, opts.WTF), {"xsi:type":"dcterms:W3CDTF"}, o, p);
		}

		for(var i = 0; i != CORE_PROPS.length; ++i) {
			var f = CORE_PROPS[i];
			var v = opts.Props && opts.Props[f[1]] != null ? opts.Props[f[1]] : cp ? cp[f[1]] : null;
			if(v === true) v = "1";
			else if(v === false) v = "0";
			else if(typeof v == "number") v = String(v);
			if(v != null) cp_doit(f[0], v, null, o, p);
		}
		if(o.length>2){ o[o.length] = ('</cp:coreProperties>'); o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}
	/* 15.2.12.3 Extended File Properties Part */
	/* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
	var EXT_PROPS/*:Array<Array<string> >*/ = [
		["Application", "Application", "string"],
		["AppVersion", "AppVersion", "string"],
		["Company", "Company", "string"],
		["DocSecurity", "DocSecurity", "string"],
		["Manager", "Manager", "string"],
		["HyperlinksChanged", "HyperlinksChanged", "bool"],
		["SharedDoc", "SharedDoc", "bool"],
		["LinksUpToDate", "LinksUpToDate", "bool"],
		["ScaleCrop", "ScaleCrop", "bool"],
		["HeadingPairs", "HeadingPairs", "raw"],
		["TitlesOfParts", "TitlesOfParts", "raw"]
	];

	function write_ext_props(cp/*::, opts*/)/*:string*/ {
		var o/*:Array<string>*/ = [], W = writextag;
		if(!cp) cp = {};
		cp.Application = "SheetJS";
		o[o.length] = (XML_HEADER);
		o[o.length] = (writextag('Properties', null, {
			'xmlns': XMLNS.EXT_PROPS,
			'xmlns:vt': XMLNS.vt
		}));

		EXT_PROPS.forEach(function(f) {
			if(cp[f[1]] === undefined) return;
			var v;
			switch(f[2]) {
				case 'string': v = escapexml(String(cp[f[1]])); break;
				case 'bool': v = cp[f[1]] ? 'true' : 'false'; break;
			}
			if(v !== undefined) o[o.length] = (W(f[0], v));
		});

		/* TODO: HeadingPairs, TitlesOfParts */
		o[o.length] = (W('HeadingPairs', W('vt:vector', W('vt:variant', '<vt:lpstr>Worksheets</vt:lpstr>')+W('vt:variant', W('vt:i4', String(cp.Worksheets))), {size:2, baseType:"variant"})));
		o[o.length] = (W('TitlesOfParts', W('vt:vector', cp.SheetNames.map(function(s) { return "<vt:lpstr>" + escapexml(s) + "</vt:lpstr>"; }).join(""), {size: cp.Worksheets, baseType:"lpstr"})));
		if(o.length>2){ o[o.length] = ('</Properties>'); o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}

	function write_cust_props(cp/*::, opts*/)/*:string*/ {
		var o = [XML_HEADER, writextag('Properties', null, {
			'xmlns': XMLNS.CUST_PROPS,
			'xmlns:vt': XMLNS.vt
		})];
		if(!cp) return o.join("");
		var pid = 1;
		keys(cp).forEach(function custprop(k) { ++pid;
			o[o.length] = (writextag('property', write_vt(cp[k], true), {
				'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
				'pid': pid,
				'name': escapexml(k)
			}));
		});
		if(o.length>2){ o[o.length] = '</Properties>'; o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}

	var straywsregex = /^\s|\s$|[\t\n\r]/;
	function write_sst_xml(sst/*:SST*/, opts)/*:string*/ {
		if(!opts.bookSST) return "";
		var o = [XML_HEADER];
		o[o.length] = (writextag('sst', null, {
			xmlns: XMLNS_main[0],
			count: sst.Count,
			uniqueCount: sst.Unique
		}));
		for(var i = 0; i != sst.length; ++i) { if(sst[i] == null) continue;
			var s/*:XLString*/ = sst[i];
			var sitag = "<si>";
			if(s.r) sitag += s.r;
			else {
				sitag += "<t";
				if(!s.t) s.t = "";
				if(s.t.match(straywsregex)) sitag += ' xml:space="preserve"';
				sitag += ">" + escapexml(s.t) + "</t>";
			}
			sitag += "</si>";
			o[o.length] = (sitag);
		}
		if(o.length>2){ o[o.length] = ('</sst>'); o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}
	function _JS2ANSI(str/*:string*/)/*:Array<number>*/ {
		var o/*:Array<number>*/ = [], oo = str.split("");
		for(var i = 0; i < oo.length; ++i) o[i] = oo[i].charCodeAt(0);
		return o;
	}

	/* [MS-OFFCRYPTO] 2.3.7.1 Binary Document Password Verifier Derivation */
	function crypto_CreatePasswordVerifier_Method1(Password/*:string*/) {
		var Verifier = 0x0000, PasswordArray;
		var PasswordDecoded = _JS2ANSI(Password);
		var len = PasswordDecoded.length + 1, i, PasswordByte;
		var Intermediate1, Intermediate2, Intermediate3;
		PasswordArray = new_raw_buf(len);
		PasswordArray[0] = PasswordDecoded.length;
		for(i = 1; i != len; ++i) PasswordArray[i] = PasswordDecoded[i-1];
		for(i = len-1; i >= 0; --i) {
			PasswordByte = PasswordArray[i];
			Intermediate1 = ((Verifier & 0x4000) === 0x0000) ? 0 : 1;
			Intermediate2 = (Verifier << 1) & 0x7FFF;
			Intermediate3 = Intermediate1 | Intermediate2;
			Verifier = Intermediate3 ^ PasswordByte;
		}
		return Verifier ^ 0xCE4B;
	}

	/* 18.3.1.13 width calculations */
	/* [MS-OI29500] 2.1.595 Column Width & Formatting */
	var DEF_MDW = 6, MDW = DEF_MDW;
	function px2char(px) { return (Math.floor((px - 5)/MDW * 100 + 0.5))/100; }
	function char2width(chr) { return (Math.round((chr * MDW + 5)/MDW*256))/256; }

	var DEF_PPI = 96, PPI = DEF_PPI;
	function px2pt(px) { return px * 96 / PPI; }

	function write_numFmts(NF/*:{[n:number|string]:string}*//*::, opts*/) {
		var o = ["<numFmts>"];
		[[5,8],[23,26],[41,44],[/*63*/50,/*66],[164,*/392]].forEach(function(r) {
			for(var i = r[0]; i <= r[1]; ++i) if(NF[i] != null) o[o.length] = (writextag('numFmt',null,{numFmtId:i,formatCode:escapexml(NF[i])}));
		});
		if(o.length === 1) return "";
		o[o.length] = ("</numFmts>");
		o[0] = writextag('numFmts', null, { count:o.length-2 }).replace("/>", ">");
		return o.join("");
	}

	function write_cellXfs(cellXfs)/*:string*/ {
		var o/*:Array<string>*/ = [];
		o[o.length] = (writextag('cellXfs',null));
		cellXfs.forEach(function(c) {
			o[o.length] = (writextag('xf', null, c));
		});
		o[o.length] = ("</cellXfs>");
		if(o.length === 2) return "";
		o[0] = writextag('cellXfs',null, {count:o.length-2}).replace("/>",">");
		return o.join("");
	}

	function write_sty_xml(wb/*:Workbook*/, opts)/*:string*/ {
		var o = [XML_HEADER, writextag('styleSheet', null, {
			'xmlns': XMLNS_main[0],
			'xmlns:vt': XMLNS.vt
		})], w;
		if(wb.SSF && (w = write_numFmts(wb.SSF)) != null) o[o.length] = w;
		o[o.length] = ('<fonts count="1"><font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>');
		o[o.length] = ('<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>');
		o[o.length] = ('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');
		o[o.length] = ('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
		if((w = write_cellXfs(opts.cellXfs))) o[o.length] = (w);
		o[o.length] = ('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
		o[o.length] = ('<dxfs count="0"/>');
		o[o.length] = ('<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>');

		if(o.length>2){ o[o.length] = ('</styleSheet>'); o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}

	function write_theme(Themes, opts)/*:string*/ {
		if(opts && opts.themeXLSX) return opts.themeXLSX;
		if(Themes && typeof Themes.raw == "string") return Themes.raw;
		var o = [XML_HEADER];
		o[o.length] = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">';
		o[o.length] =  '<a:themeElements>';

		o[o.length] =   '<a:clrScheme name="Office">';
		o[o.length] =    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>';
		o[o.length] =    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>';
		o[o.length] =    '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>';
		o[o.length] =    '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>';
		o[o.length] =    '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>';
		o[o.length] =    '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>';
		o[o.length] =    '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>';
		o[o.length] =    '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>';
		o[o.length] =    '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>';
		o[o.length] =    '<a:accent6><a:srgbClr val="F79646"/></a:accent6>';
		o[o.length] =    '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>';
		o[o.length] =    '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>';
		o[o.length] =   '</a:clrScheme>';

		o[o.length] =   '<a:fontScheme name="Office">';
		o[o.length] =    '<a:majorFont>';
		o[o.length] =     '<a:latin typeface="Cambria"/>';
		o[o.length] =     '<a:ea typeface=""/>';
		o[o.length] =     '<a:cs typeface=""/>';
		o[o.length] =     '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
		o[o.length] =     '<a:font script="Hang" typeface="맑은 고딕"/>';
		o[o.length] =     '<a:font script="Hans" typeface="宋体"/>';
		o[o.length] =     '<a:font script="Hant" typeface="新細明體"/>';
		o[o.length] =     '<a:font script="Arab" typeface="Times New Roman"/>';
		o[o.length] =     '<a:font script="Hebr" typeface="Times New Roman"/>';
		o[o.length] =     '<a:font script="Thai" typeface="Tahoma"/>';
		o[o.length] =     '<a:font script="Ethi" typeface="Nyala"/>';
		o[o.length] =     '<a:font script="Beng" typeface="Vrinda"/>';
		o[o.length] =     '<a:font script="Gujr" typeface="Shruti"/>';
		o[o.length] =     '<a:font script="Khmr" typeface="MoolBoran"/>';
		o[o.length] =     '<a:font script="Knda" typeface="Tunga"/>';
		o[o.length] =     '<a:font script="Guru" typeface="Raavi"/>';
		o[o.length] =     '<a:font script="Cans" typeface="Euphemia"/>';
		o[o.length] =     '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
		o[o.length] =     '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
		o[o.length] =     '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
		o[o.length] =     '<a:font script="Thaa" typeface="MV Boli"/>';
		o[o.length] =     '<a:font script="Deva" typeface="Mangal"/>';
		o[o.length] =     '<a:font script="Telu" typeface="Gautami"/>';
		o[o.length] =     '<a:font script="Taml" typeface="Latha"/>';
		o[o.length] =     '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
		o[o.length] =     '<a:font script="Orya" typeface="Kalinga"/>';
		o[o.length] =     '<a:font script="Mlym" typeface="Kartika"/>';
		o[o.length] =     '<a:font script="Laoo" typeface="DokChampa"/>';
		o[o.length] =     '<a:font script="Sinh" typeface="Iskoola Pota"/>';
		o[o.length] =     '<a:font script="Mong" typeface="Mongolian Baiti"/>';
		o[o.length] =     '<a:font script="Viet" typeface="Times New Roman"/>';
		o[o.length] =     '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
		o[o.length] =     '<a:font script="Geor" typeface="Sylfaen"/>';
		o[o.length] =    '</a:majorFont>';
		o[o.length] =    '<a:minorFont>';
		o[o.length] =     '<a:latin typeface="Calibri"/>';
		o[o.length] =     '<a:ea typeface=""/>';
		o[o.length] =     '<a:cs typeface=""/>';
		o[o.length] =     '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
		o[o.length] =     '<a:font script="Hang" typeface="맑은 고딕"/>';
		o[o.length] =     '<a:font script="Hans" typeface="宋体"/>';
		o[o.length] =     '<a:font script="Hant" typeface="新細明體"/>';
		o[o.length] =     '<a:font script="Arab" typeface="Arial"/>';
		o[o.length] =     '<a:font script="Hebr" typeface="Arial"/>';
		o[o.length] =     '<a:font script="Thai" typeface="Tahoma"/>';
		o[o.length] =     '<a:font script="Ethi" typeface="Nyala"/>';
		o[o.length] =     '<a:font script="Beng" typeface="Vrinda"/>';
		o[o.length] =     '<a:font script="Gujr" typeface="Shruti"/>';
		o[o.length] =     '<a:font script="Khmr" typeface="DaunPenh"/>';
		o[o.length] =     '<a:font script="Knda" typeface="Tunga"/>';
		o[o.length] =     '<a:font script="Guru" typeface="Raavi"/>';
		o[o.length] =     '<a:font script="Cans" typeface="Euphemia"/>';
		o[o.length] =     '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
		o[o.length] =     '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
		o[o.length] =     '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
		o[o.length] =     '<a:font script="Thaa" typeface="MV Boli"/>';
		o[o.length] =     '<a:font script="Deva" typeface="Mangal"/>';
		o[o.length] =     '<a:font script="Telu" typeface="Gautami"/>';
		o[o.length] =     '<a:font script="Taml" typeface="Latha"/>';
		o[o.length] =     '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
		o[o.length] =     '<a:font script="Orya" typeface="Kalinga"/>';
		o[o.length] =     '<a:font script="Mlym" typeface="Kartika"/>';
		o[o.length] =     '<a:font script="Laoo" typeface="DokChampa"/>';
		o[o.length] =     '<a:font script="Sinh" typeface="Iskoola Pota"/>';
		o[o.length] =     '<a:font script="Mong" typeface="Mongolian Baiti"/>';
		o[o.length] =     '<a:font script="Viet" typeface="Arial"/>';
		o[o.length] =     '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
		o[o.length] =     '<a:font script="Geor" typeface="Sylfaen"/>';
		o[o.length] =    '</a:minorFont>';
		o[o.length] =   '</a:fontScheme>';

		o[o.length] =   '<a:fmtScheme name="Office">';
		o[o.length] =    '<a:fillStyleLst>';
		o[o.length] =     '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
		o[o.length] =     '<a:gradFill rotWithShape="1">';
		o[o.length] =      '<a:gsLst>';
		o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		o[o.length] =      '</a:gsLst>';
		o[o.length] =      '<a:lin ang="16200000" scaled="1"/>';
		o[o.length] =     '</a:gradFill>';
		o[o.length] =     '<a:gradFill rotWithShape="1">';
		o[o.length] =      '<a:gsLst>';
		o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		o[o.length] =      '</a:gsLst>';
		o[o.length] =      '<a:lin ang="16200000" scaled="0"/>';
		o[o.length] =     '</a:gradFill>';
		o[o.length] =    '</a:fillStyleLst>';
		o[o.length] =    '<a:lnStyleLst>';
		o[o.length] =     '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>';
		o[o.length] =     '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
		o[o.length] =     '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
		o[o.length] =    '</a:lnStyleLst>';
		o[o.length] =    '<a:effectStyleLst>';
		o[o.length] =     '<a:effectStyle>';
		o[o.length] =      '<a:effectLst>';
		o[o.length] =       '<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw>';
		o[o.length] =      '</a:effectLst>';
		o[o.length] =     '</a:effectStyle>';
		o[o.length] =     '<a:effectStyle>';
		o[o.length] =      '<a:effectLst>';
		o[o.length] =       '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
		o[o.length] =      '</a:effectLst>';
		o[o.length] =     '</a:effectStyle>';
		o[o.length] =     '<a:effectStyle>';
		o[o.length] =      '<a:effectLst>';
		o[o.length] =       '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
		o[o.length] =      '</a:effectLst>';
		o[o.length] =      '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d>';
		o[o.length] =      '<a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>';
		o[o.length] =     '</a:effectStyle>';
		o[o.length] =    '</a:effectStyleLst>';
		o[o.length] =    '<a:bgFillStyleLst>';
		o[o.length] =     '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
		o[o.length] =     '<a:gradFill rotWithShape="1">';
		o[o.length] =      '<a:gsLst>';
		o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>';
		o[o.length] =      '</a:gsLst>';
		o[o.length] =      '<a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>';
		o[o.length] =     '</a:gradFill>';
		o[o.length] =     '<a:gradFill rotWithShape="1">';
		o[o.length] =      '<a:gsLst>';
		o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
		o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>';
		o[o.length] =      '</a:gsLst>';
		o[o.length] =      '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>';
		o[o.length] =     '</a:gradFill>';
		o[o.length] =    '</a:bgFillStyleLst>';
		o[o.length] =   '</a:fmtScheme>';
		o[o.length] =  '</a:themeElements>';

		o[o.length] =  '<a:objectDefaults>';
		o[o.length] =   '<a:spDef>';
		o[o.length] =    '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style>';
		o[o.length] =   '</a:spDef>';
		o[o.length] =   '<a:lnDef>';
		o[o.length] =    '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style>';
		o[o.length] =   '</a:lnDef>';
		o[o.length] =  '</a:objectDefaults>';
		o[o.length] =  '<a:extraClrSchemeLst/>';
		o[o.length] = '</a:theme>';
		return o.join("");
	}
	function write_xlmeta_xml() {
	  var o = [XML_HEADER];
	  o.push('<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">\n  <metadataTypes count="1">\n    <metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>\n  </metadataTypes>\n  <futureMetadata name="XLDAPR" count="1">\n    <bk>\n      <extLst>\n        <ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">\n          <xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>\n        </ext>\n      </extLst>\n    </bk>\n  </futureMetadata>\n  <cellMetadata count="1">\n    <bk>\n      <rc t="1" v="0"/>\n    </bk>\n  </cellMetadata>\n</metadata>');
	  return o.join("");
	}

	/* L.5.5.2 SpreadsheetML Comments + VML Schema */
	var _shapeid = 1024;
	function write_comments_vml(rId/*:number*/, comments) {
		var csize = [21600, 21600];
		/* L.5.2.1.2 Path Attribute */
		var bbox = ["m0,0l0",csize[1],csize[0],csize[1],csize[0],"0xe"].join(",");
		var o = [
			writextag("xml", null, { 'xmlns:v': XLMLNS.v, 'xmlns:o': XLMLNS.o, 'xmlns:x': XLMLNS.x, 'xmlns:mv': XLMLNS.mv }).replace(/\/>/,">"),
			writextag("o:shapelayout", writextag("o:idmap", null, {'v:ext':"edit", 'data':rId}), {'v:ext':"edit"}),
			writextag("v:shapetype", [
				writextag("v:stroke", null, {joinstyle:"miter"}),
				writextag("v:path", null, {gradientshapeok:"t", 'o:connecttype':"rect"})
			].join(""), {id:"_x0000_t202", 'o:spt':202, coordsize:csize.join(","),path:bbox})
		];
		while(_shapeid < rId * 1000) _shapeid += 1000;

		comments.forEach(function(x) {
		var c = decode_cell(x[0]);
		var fillopts = /*::(*/{'color2':"#BEFF82", 'type':"gradient"}/*:: :any)*/;
		if(fillopts.type == "gradient") fillopts.angle = "-180";
		var fillparm = fillopts.type == "gradient" ? writextag("o:fill", null, {type:"gradientUnscaled", 'v:ext':"view"}) : null;
		var fillxml = writextag('v:fill', fillparm, fillopts);

		var shadata = ({on:"t", 'obscured':"t"}/*:any*/);
		++_shapeid;

		o = o.concat([
		'<v:shape' + wxt_helper({
			id:'_x0000_s' + _shapeid,
			type:"#_x0000_t202",
			style:"position:absolute; margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:10" + (x[1].hidden ? ";visibility:hidden" : "") ,
			fillcolor:"#ECFAD4",
			strokecolor:"#edeaa1"
		}) + '>',
			fillxml,
			writextag("v:shadow", null, shadata),
			writextag("v:path", null, {'o:connecttype':"none"}),
			'<v:textbox><div style="text-align:left"></div></v:textbox>',
			'<x:ClientData ObjectType="Note">',
				'<x:MoveWithCells/>',
				'<x:SizeWithCells/>',
				/* Part 4 19.4.2.3 Anchor (Anchor) */
				writetag('x:Anchor', [c.c+1, 0, c.r+1, 0, c.c+3, 20, c.r+5, 20].join(",")),
				writetag('x:AutoFill', "False"),
				writetag('x:Row', String(c.r)),
				writetag('x:Column', String(c.c)),
				x[1].hidden ? '' : '<x:Visible/>',
			'</x:ClientData>',
		'</v:shape>'
		]); });
		o.push('</xml>');
		return o.join("");
	}

	function write_comments_xml(data/*::, opts*/) {
		var o = [XML_HEADER, writextag('comments', null, { 'xmlns': XMLNS_main[0] })];

		var iauthor/*:Array<string>*/ = [];
		o.push("<authors>");
		data.forEach(function(x) { x[1].forEach(function(w) { var a = escapexml(w.a);
			if(iauthor.indexOf(a) == -1) {
				iauthor.push(a);
				o.push("<author>" + a + "</author>");
			}
			if(w.T && w.ID && iauthor.indexOf("tc=" + w.ID) == -1) {
				iauthor.push("tc=" + w.ID);
				o.push("<author>" + "tc=" + w.ID + "</author>");
			}
		}); });
		if(iauthor.length == 0) { iauthor.push("SheetJ5"); o.push("<author>SheetJ5</author>"); }
		o.push("</authors>");
		o.push("<commentList>");
		data.forEach(function(d) {
			/* 18.7.3 CT_Comment */
			var lastauthor = 0, ts = [];
			if(d[1][0] && d[1][0].T && d[1][0].ID) lastauthor = iauthor.indexOf("tc=" + d[1][0].ID);
			else d[1].forEach(function(c) {
				if(c.a) lastauthor = iauthor.indexOf(escapexml(c.a));
				ts.push(c.t||"");
			});
			o.push('<comment ref="' + d[0] + '" authorId="' + lastauthor + '"><text>');
			if(ts.length <= 1) o.push(writetag("t", escapexml(ts[0]||"")));
			else {
				/* based on Threaded Comments -> Comments projection */
				var t = "Comment:\n    " + (ts[0]) + "\n";
				for(var i = 1; i < ts.length; ++i) t += "Reply:\n    " + ts[i] + "\n";
				o.push(writetag("t", escapexml(t)));
			}
			o.push('</text></comment>');
		});
		o.push("</commentList>");
		if(o.length>2) { o[o.length] = ('</comments>'); o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}

	function write_tcmnt_xml(comments, people, opts) {
		var o = [XML_HEADER, writextag('ThreadedComments', null, { 'xmlns': XMLNS.TCMNT }).replace(/[\/]>/, ">")];
		comments.forEach(function(carr) {
			var rootid = "";
			(carr[1] || []).forEach(function(c, idx) {
				if(!c.T) { delete c.ID; return; }
				if(c.a && people.indexOf(c.a) == -1) people.push(c.a);
				var tcopts = {
					ref: carr[0],
					id: "{54EE7951-7262-4200-6969-" + ("000000000000" + opts.tcid++).slice(-12) + "}"
				};
				if(idx == 0) rootid = tcopts.id;
				else tcopts.parentId = rootid;
				c.ID = tcopts.id;
				if(c.a) tcopts.personId = "{54EE7950-7262-4200-6969-" + ("000000000000" + people.indexOf(c.a)).slice(-12) + "}";
				o.push(writextag('threadedComment', writetag('text', c.t||""), tcopts));
			});
		});
		o.push('</ThreadedComments>');
		return o.join("");
	}
	function write_people_xml(people/*, opts*/) {
		var o = [XML_HEADER, writextag('personList', null, {
			'xmlns': XMLNS.TCMNT,
			'xmlns:x': XMLNS_main[0]
		}).replace(/[\/]>/, ">")];
		people.forEach(function(person, idx) {
			o.push(writextag('person', null, {
				displayName: person,
				id: "{54EE7950-7262-4200-6969-" + ("000000000000" + idx).slice(-12) + "}",
				userId: person,
				providerId: "None"
			}));
		});
		o.push("</personList>");
		return o.join("");
	}
	var VBAFMTS = ["xlsb", "xlsm", "xlam", "biff8", "xla"];


	/*global Map */
	var browser_has_Map = typeof Map !== 'undefined';

	function get_sst_id(sst/*:SST*/, str/*:string*/, rev)/*:number*/ {
		var i = 0, len = sst.length;
		if(rev) {
			if(browser_has_Map ? rev.has(str) : Object.prototype.hasOwnProperty.call(rev, str)) {
				var revarr = browser_has_Map ? rev.get(str) : rev[str];
				for(; i < revarr.length; ++i) {
					if(sst[revarr[i]].t === str) { sst.Count ++; return revarr[i]; }
				}
			}
		} else for(; i < len; ++i) {
			if(sst[i].t === str) { sst.Count ++; return i; }
		}
		sst[len] = ({t:str}/*:any*/); sst.Count ++; sst.Unique ++;
		if(rev) {
			if(browser_has_Map) {
				if(!rev.has(str)) rev.set(str, []);
				rev.get(str).push(len);
			} else {
				if(!Object.prototype.hasOwnProperty.call(rev, str)) rev[str] = [];
				rev[str].push(len);
			}
		}
		return len;
	}

	function col_obj_w(C/*:number*/, col) {
		var p = ({min:C+1,max:C+1}/*:any*/);
		/* wch (chars), wpx (pixels) */
		var wch = -1;
		if(col.MDW) MDW = col.MDW;
		if(col.width != null) p.customWidth = 1;
		else if(col.wpx != null) wch = px2char(col.wpx);
		else if(col.wch != null) wch = col.wch;
		if(wch > -1) { p.width = char2width(wch); p.customWidth = 1; }
		else if(col.width != null) p.width = col.width;
		if(col.hidden) p.hidden = true;
		if(col.level != null) { p.outlineLevel = p.level = col.level; }
		return p;
	}

	function default_margins(margins/*:Margins*/, mode/*:?string*/) {
		if(!margins) return;
		var defs = [0.7, 0.7, 0.75, 0.75, 0.3, 0.3];
		if(mode == 'xlml') defs = [1, 1, 1, 1, 0.5, 0.5];
		if(margins.left   == null) margins.left   = defs[0];
		if(margins.right  == null) margins.right  = defs[1];
		if(margins.top    == null) margins.top    = defs[2];
		if(margins.bottom == null) margins.bottom = defs[3];
		if(margins.header == null) margins.header = defs[4];
		if(margins.footer == null) margins.footer = defs[5];
	}

	function get_cell_style(styles/*:Array<any>*/, cell/*:Cell*/, opts) {
		var z = opts.revssf[cell.z != null ? cell.z : "General"];
		var i = 0x3c, len = styles.length;
		if(z == null && opts.ssf) {
			for(; i < 0x188; ++i) if(opts.ssf[i] == null) {
				SSF_load(cell.z, i);
				// $FlowIgnore
				opts.ssf[i] = cell.z;
				opts.revssf[cell.z] = z = i;
				break;
			}
		}
		for(i = 0; i != len; ++i) if(styles[i].numFmtId === z) return i;
		styles[len] = {
			numFmtId:z,
			fontId:0,
			fillId:0,
			borderId:0,
			xfId:0,
			applyNumberFormat:1
		};
		return len;
	}

	function check_ws(ws/*:Worksheet*/, sname/*:string*/, i/*:number*/) {
		if(ws && ws['!ref']) {
			var range = safe_decode_range(ws['!ref']);
			if(range.e.c < range.s.c || range.e.r < range.s.r) throw new Error("Bad range (" + i + "): " + ws['!ref']);
		}
	}

	function write_ws_xml_merges(merges/*:Array<Range>*/)/*:string*/ {
		if(merges.length === 0) return "";
		var o = '<mergeCells count="' + merges.length + '">';
		for(var i = 0; i != merges.length; ++i) o += '<mergeCell ref="' + encode_range(merges[i]) + '"/>';
		return o + '</mergeCells>';
	}
	function write_ws_xml_sheetpr(ws, wb, idx, opts, o) {
		var needed = false;
		var props = {}, payload = null;
		if(opts.bookType !== 'xlsx' && wb.vbaraw) {
			var cname = wb.SheetNames[idx];
			try { if(wb.Workbook) cname = wb.Workbook.Sheets[idx].CodeName || cname; } catch(e) {}
			needed = true;
			props.codeName = utf8write(escapexml(cname));
		}

		if(ws && ws["!outline"]) {
			var outlineprops = {summaryBelow:1, summaryRight:1};
			if(ws["!outline"].above) outlineprops.summaryBelow = 0;
			if(ws["!outline"].left) outlineprops.summaryRight = 0;
			payload = (payload||"") + writextag('outlinePr', null, outlineprops);
		}

		if(!needed && !payload) return;
		o[o.length] = (writextag('sheetPr', payload, props));
	}

	/* 18.3.1.85 sheetProtection CT_SheetProtection */
	var sheetprot_deffalse = ["objects", "scenarios", "selectLockedCells", "selectUnlockedCells"];
	var sheetprot_deftrue = [
		"formatColumns", "formatRows", "formatCells",
		"insertColumns", "insertRows", "insertHyperlinks",
		"deleteColumns", "deleteRows",
		"sort", "autoFilter", "pivotTables"
	];
	function write_ws_xml_protection(sp)/*:string*/ {
		// algorithmName, hashValue, saltValue, spinCount
		var o = ({sheet:1}/*:any*/);
		sheetprot_deffalse.forEach(function(n) { if(sp[n] != null && sp[n]) o[n] = "1"; });
		sheetprot_deftrue.forEach(function(n) { if(sp[n] != null && !sp[n]) o[n] = "0"; });
		/* TODO: algorithm */
		if(sp.password) o.password = crypto_CreatePasswordVerifier_Method1(sp.password).toString(16).toUpperCase();
		return writextag('sheetProtection', null, o);
	}
	function write_ws_xml_margins(margin)/*:string*/ {
		default_margins(margin);
		return writextag('pageMargins', null, margin);
	}
	function write_ws_xml_cols(ws, cols)/*:string*/ {
		var o = ["<cols>"], col;
		for(var i = 0; i != cols.length; ++i) {
			if(!(col = cols[i])) continue;
			o[o.length] = (writextag('col', null, col_obj_w(i, col)));
		}
		o[o.length] = "</cols>";
		return o.join("");
	}
	function write_ws_xml_autofilter(data, ws, wb, idx)/*:string*/ {
		var ref = typeof data.ref == "string" ? data.ref : encode_range(data.ref);
		if(!wb.Workbook) wb.Workbook = ({Sheets:[]}/*:any*/);
		if(!wb.Workbook.Names) wb.Workbook.Names = [];
		var names/*: Array<any> */ = wb.Workbook.Names;
		var range = decode_range(ref);
		if(range.s.r == range.e.r) { range.e.r = decode_range(ws["!ref"]).e.r; ref = encode_range(range); }
		for(var i = 0; i < names.length; ++i) {
			var name = names[i];
			if(name.Name != '_xlnm._FilterDatabase') continue;
			if(name.Sheet != idx) continue;
			name.Ref = "'" + wb.SheetNames[idx] + "'!" + ref; break;
		}
		if(i == names.length) names.push({ Name: '_xlnm._FilterDatabase', Sheet: idx, Ref: "'" + wb.SheetNames[idx] + "'!" + ref  });
		return writextag("autoFilter", null, {ref:ref});
	}
	function write_ws_xml_sheetviews(ws, opts, idx, wb)/*:string*/ {
		var sview = ({workbookViewId:"0"}/*:any*/);
		// $FlowIgnore
		if((((wb||{}).Workbook||{}).Views||[])[0]) sview.rightToLeft = wb.Workbook.Views[0].RTL ? "1" : "0";
		return writextag("sheetViews", writextag("sheetView", null, sview), {});
	}

	function write_ws_xml_cell(cell/*:Cell*/, ref, ws, opts/*::, idx, wb*/)/*:string*/ {
		if(cell.c) ws['!comments'].push([ref, cell.c]);
		if(cell.v === undefined && typeof cell.f !== "string" || cell.t === 'z' && !cell.f) return "";
		var vv = "";
		var oldt = cell.t, oldv = cell.v;
		if(cell.t !== "z") switch(cell.t) {
			case 'b': vv = cell.v ? "1" : "0"; break;
			case 'n': vv = ''+cell.v; break;
			case 'e': vv = BErr[cell.v]; break;
			case 'd':
				if(opts && opts.cellDates) vv = parseDate(cell.v, -1).toISOString();
				else {
					cell = dup(cell);
					cell.t = 'n';
					vv = ''+(cell.v = datenum(parseDate(cell.v)));
				}
				if(typeof cell.z === 'undefined') cell.z = table_fmt[14];
				break;
			default: vv = cell.v; break;
		}
		var v = writetag('v', escapexml(vv)), o = ({r:ref}/*:any*/);
		/* TODO: cell style */
		var os = get_cell_style(opts.cellXfs, cell, opts);
		if(os !== 0) o.s = os;
		switch(cell.t) {
			case 'n': break;
			case 'd': o.t = "d"; break;
			case 'b': o.t = "b"; break;
			case 'e': o.t = "e"; break;
			case 'z': break;
			default: if(cell.v == null) { delete cell.t; break; }
				if(cell.v.length > 32767) throw new Error("Text length must not exceed 32767 characters");
				if(opts && opts.bookSST) {
					v = writetag('v', ''+get_sst_id(opts.Strings, cell.v, opts.revStrings));
					o.t = "s"; break;
				}
				o.t = "str"; break;
		}
		if(cell.t != oldt) { cell.t = oldt; cell.v = oldv; }
		if(typeof cell.f == "string" && cell.f) {
			var ff = cell.F && cell.F.slice(0, ref.length) == ref ? {t:"array", ref:cell.F} : null;
			v = writextag('f', escapexml(cell.f), ff) + (cell.v != null ? v : "");
		}
		if(cell.l) ws['!links'].push([ref, cell.l]);
		if(cell.D) o.cm = 1;
		return writextag('c', v, o);
	}

	function write_ws_xml_data(ws/*:Worksheet*/, opts, idx/*:number*/, wb/*:Workbook*//*::, rels*/)/*:string*/ {
		var o/*:Array<string>*/ = [], r/*:Array<string>*/ = [], range = safe_decode_range(ws['!ref']), cell="", ref, rr = "", cols/*:Array<string>*/ = [], R=0, C=0, rows = ws['!rows'];
		var dense = Array.isArray(ws);
		var params = ({r:rr}/*:any*/), row/*:RowInfo*/, height = -1;
		for(C = range.s.c; C <= range.e.c; ++C) cols[C] = encode_col(C);
		for(R = range.s.r; R <= range.e.r; ++R) {
			r = [];
			rr = encode_row(R);
			for(C = range.s.c; C <= range.e.c; ++C) {
				ref = cols[C] + rr;
				var _cell = dense ? (ws[R]||[])[C]: ws[ref];
				if(_cell === undefined) continue;
				if((cell = write_ws_xml_cell(_cell, ref, ws, opts)) != null) r.push(cell);
			}
			if(r.length > 0 || (rows && rows[R])) {
				params = ({r:rr}/*:any*/);
				if(rows && rows[R]) {
					row = rows[R];
					if(row.hidden) params.hidden = 1;
					height = -1;
					if(row.hpx) height = px2pt(row.hpx);
					else if(row.hpt) height = row.hpt;
					if(height > -1) { params.ht = height; params.customHeight = 1; }
					if(row.level) { params.outlineLevel = row.level; }
				}
				o[o.length] = (writextag('row', r.join(""), params));
			}
		}
		if(rows) for(; R < rows.length; ++R) {
			if(rows && rows[R]) {
				params = ({r:R+1}/*:any*/);
				row = rows[R];
				if(row.hidden) params.hidden = 1;
				height = -1;
				if (row.hpx) height = px2pt(row.hpx);
				else if (row.hpt) height = row.hpt;
				if (height > -1) { params.ht = height; params.customHeight = 1; }
				if (row.level) { params.outlineLevel = row.level; }
				o[o.length] = (writextag('row', "", params));
			}
		}
		return o.join("");
	}

	function write_ws_xml(idx/*:number*/, opts, wb/*:Workbook*/, rels)/*:string*/ {
		var o = [XML_HEADER, writextag('worksheet', null, {
			'xmlns': XMLNS_main[0],
			'xmlns:r': XMLNS.r
		})];
		var s = wb.SheetNames[idx], sidx = 0, rdata = "";
		var ws = wb.Sheets[s];
		if(ws == null) ws = {};
		var ref = ws['!ref'] || 'A1';
		var range = safe_decode_range(ref);
		if(range.e.c > 0x3FFF || range.e.r > 0xFFFFF) {
			if(opts.WTF) throw new Error("Range " + ref + " exceeds format limit A1:XFD1048576");
			range.e.c = Math.min(range.e.c, 0x3FFF);
			range.e.r = Math.min(range.e.c, 0xFFFFF);
			ref = encode_range(range);
		}
		if(!rels) rels = {};
		ws['!comments'] = [];
		var _drawing = [];

		write_ws_xml_sheetpr(ws, wb, idx, opts, o);

		o[o.length] = (writextag('dimension', null, {'ref': ref}));

		o[o.length] = write_ws_xml_sheetviews(ws, opts, idx, wb);

		/* TODO: store in WB, process styles */
		if(opts.sheetFormat) o[o.length] = (writextag('sheetFormatPr', null, {
			defaultRowHeight:opts.sheetFormat.defaultRowHeight||'16',
			baseColWidth:opts.sheetFormat.baseColWidth||'10',
			outlineLevelRow:opts.sheetFormat.outlineLevelRow||'7'
		}));

		if(ws['!cols'] != null && ws['!cols'].length > 0) o[o.length] = (write_ws_xml_cols(ws, ws['!cols']));

		o[sidx = o.length] = '<sheetData/>';
		ws['!links'] = [];
		if(ws['!ref'] != null) {
			rdata = write_ws_xml_data(ws, opts);
			if(rdata.length > 0) o[o.length] = (rdata);
		}
		if(o.length>sidx+1) { o[o.length] = ('</sheetData>'); o[sidx]=o[sidx].replace("/>",">"); }

		/* sheetCalcPr */

		if(ws['!protect']) o[o.length] = write_ws_xml_protection(ws['!protect']);

		/* protectedRanges */
		/* scenarios */

		if(ws['!autofilter'] != null) o[o.length] = write_ws_xml_autofilter(ws['!autofilter'], ws, wb, idx);

		/* sortState */
		/* dataConsolidate */
		/* customSheetViews */

		if(ws['!merges'] != null && ws['!merges'].length > 0) o[o.length] = (write_ws_xml_merges(ws['!merges']));

		/* phoneticPr */
		/* conditionalFormatting */
		/* dataValidations */

		var relc = -1, rel, rId = -1;
		if(/*::(*/ws['!links']/*::||[])*/.length > 0) {
			o[o.length] = "<hyperlinks>";
			/*::(*/ws['!links']/*::||[])*/.forEach(function(l) {
				if(!l[1].Target) return;
				rel = ({"ref":l[0]}/*:any*/);
				if(l[1].Target.charAt(0) != "#") {
					rId = add_rels(rels, -1, escapexml(l[1].Target).replace(/#.*$/, ""), RELS.HLINK);
					rel["r:id"] = "rId"+rId;
				}
				if((relc = l[1].Target.indexOf("#")) > -1) rel.location = escapexml(l[1].Target.slice(relc+1));
				if(l[1].Tooltip) rel.tooltip = escapexml(l[1].Tooltip);
				o[o.length] = writextag("hyperlink",null,rel);
			});
			o[o.length] = "</hyperlinks>";
		}
		delete ws['!links'];

		/* printOptions */

		if(ws['!margins'] != null) o[o.length] =  write_ws_xml_margins(ws['!margins']);

		/* pageSetup */
		/* headerFooter */
		/* rowBreaks */
		/* colBreaks */
		/* customProperties */
		/* cellWatches */

		if(!opts || opts.ignoreEC || (opts.ignoreEC == (void 0))) o[o.length] = writetag("ignoredErrors", writextag("ignoredError", null, {numberStoredAsText:1, sqref:ref}));

		/* smartTags */

		if(_drawing.length > 0) {
			rId = add_rels(rels, -1, "../drawings/drawing" + (idx+1) + ".xml", RELS.DRAW);
			o[o.length] = writextag("drawing", null, {"r:id":"rId" + rId});
			ws['!drawing'] = _drawing;
		}

		if(ws['!comments'].length > 0) {
			rId = add_rels(rels, -1, "../drawings/vmlDrawing" + (idx+1) + ".vml", RELS.VML);
			o[o.length] = writextag("legacyDrawing", null, {"r:id":"rId" + rId});
			ws['!legacy'] = rId;
		}

		/* legacyDrawingHF */
		/* picture */
		/* oleObjects */
		/* controls */
		/* webPublishItems */
		/* tableParts */
		/* extLst */

		if(o.length>1) { o[o.length] = ('</worksheet>'); o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}
	/* 18.2.28 (CT_WorkbookProtection) Defaults */
	var WBPropsDef = [
		['allowRefreshQuery',           false, "bool"],
		['autoCompressPictures',        true,  "bool"],
		['backupFile',                  false, "bool"],
		['checkCompatibility',          false, "bool"],
		['CodeName',                    ''],
		['date1904',                    false, "bool"],
		['defaultThemeVersion',         0,      "int"],
		['filterPrivacy',               false, "bool"],
		['hidePivotFieldList',          false, "bool"],
		['promptedSolutions',           false, "bool"],
		['publishItems',                false, "bool"],
		['refreshAllConnections',       false, "bool"],
		['saveExternalLinkValues',      true,  "bool"],
		['showBorderUnselectedTables',  true,  "bool"],
		['showInkAnnotation',           true,  "bool"],
		['showObjects',                 'all'],
		['showPivotChartFilter',        false, "bool"],
		['updateLinks', 'userSet']
	];

	var badchars = /*#__PURE__*/"][*?\/\\".split("");
	function check_ws_name(n/*:string*/, safe/*:?boolean*/)/*:boolean*/ {
		if(n.length > 31) { if(safe) return false; throw new Error("Sheet names cannot exceed 31 chars"); }
		var _good = true;
		badchars.forEach(function(c) {
			if(n.indexOf(c) == -1) return;
			if(!safe) throw new Error("Sheet name cannot contain : \\ / ? * [ ]");
			_good = false;
		});
		return _good;
	}
	function check_wb_names(N, S, codes) {
		N.forEach(function(n,i) {
			check_ws_name(n);
			for(var j = 0; j < i; ++j) if(n == N[j]) throw new Error("Duplicate Sheet Name: " + n);
			if(codes) {
				var cn = (S && S[i] && S[i].CodeName) || n;
				if(cn.charCodeAt(0) == 95 && cn.length > 22) throw new Error("Bad Code Name: Worksheet" + cn);
			}
		});
	}
	function check_wb(wb) {
		if(!wb || !wb.SheetNames || !wb.Sheets) throw new Error("Invalid Workbook");
		if(!wb.SheetNames.length) throw new Error("Workbook is empty");
		var Sheets = (wb.Workbook && wb.Workbook.Sheets) || [];
		check_wb_names(wb.SheetNames, Sheets, !!wb.vbaraw);
		for(var i = 0; i < wb.SheetNames.length; ++i) check_ws(wb.Sheets[wb.SheetNames[i]], wb.SheetNames[i], i);
		/* TODO: validate workbook */
	}

	function write_wb_xml(wb/*:Workbook*//*::, opts:?WriteOpts*/)/*:string*/ {
		var o = [XML_HEADER];
		o[o.length] = writextag('workbook', null, {
			'xmlns': XMLNS_main[0],
			//'xmlns:mx': XMLNS.mx,
			//'xmlns:s': XMLNS_main[0],
			'xmlns:r': XMLNS.r
		});

		var write_names = (wb.Workbook && (wb.Workbook.Names||[]).length > 0);

		/* fileVersion */
		/* fileSharing */

		var workbookPr/*:any*/ = ({codeName:"ThisWorkbook"}/*:any*/);
		if(wb.Workbook && wb.Workbook.WBProps) {
			WBPropsDef.forEach(function(x) {
				/*:: if(!wb.Workbook || !wb.Workbook.WBProps) throw "unreachable"; */
				if((wb.Workbook.WBProps[x[0]]/*:any*/) == null) return;
				if((wb.Workbook.WBProps[x[0]]/*:any*/) == x[1]) return;
				workbookPr[x[0]] = (wb.Workbook.WBProps[x[0]]/*:any*/);
			});
			/*:: if(!wb.Workbook || !wb.Workbook.WBProps) throw "unreachable"; */
			if(wb.Workbook.WBProps.CodeName) { workbookPr.codeName = wb.Workbook.WBProps.CodeName; delete workbookPr.CodeName; }
		}
		o[o.length] = (writextag('workbookPr', null, workbookPr));

		/* workbookProtection */

		var sheets = wb.Workbook && wb.Workbook.Sheets || [];
		var i = 0;

		/* bookViews only written if first worksheet is hidden */
		if(sheets && sheets[0] && !!sheets[0].Hidden) {
			o[o.length] = "<bookViews>";
			for(i = 0; i != wb.SheetNames.length; ++i) {
				if(!sheets[i]) break;
				if(!sheets[i].Hidden) break;
			}
			if(i == wb.SheetNames.length) i = 0;
			o[o.length] = '<workbookView firstSheet="' + i + '" activeTab="' + i + '"/>';
			o[o.length] = "</bookViews>";
		}

		o[o.length] = "<sheets>";
		for(i = 0; i != wb.SheetNames.length; ++i) {
			var sht = ({name:escapexml(wb.SheetNames[i].slice(0,31))}/*:any*/);
			sht.sheetId = ""+(i+1);
			sht["r:id"] = "rId"+(i+1);
			if(sheets[i]) switch(sheets[i].Hidden) {
				case 1: sht.state = "hidden"; break;
				case 2: sht.state = "veryHidden"; break;
			}
			o[o.length] = (writextag('sheet',null,sht));
		}
		o[o.length] = "</sheets>";

		/* functionGroups */
		/* externalReferences */

		if(write_names) {
			o[o.length] = "<definedNames>";
			if(wb.Workbook && wb.Workbook.Names) wb.Workbook.Names.forEach(function(n) {
				var d/*:any*/ = {name:n.Name};
				if(n.Comment) d.comment = n.Comment;
				if(n.Sheet != null) d.localSheetId = ""+n.Sheet;
				if(n.Hidden) d.hidden = "1";
				if(!n.Ref) return;
				o[o.length] = writextag('definedName', escapexml(n.Ref), d);
			});
			o[o.length] = "</definedNames>";
		}

		/* calcPr */
		/* oleSize */
		/* customWorkbookViews */
		/* pivotCaches */
		/* smartTagPr */
		/* smartTagTypes */
		/* webPublishing */
		/* fileRecoveryPr */
		/* webPublishObjects */
		/* extLst */

		if(o.length>2){ o[o.length] = '</workbook>'; o[1]=o[1].replace("/>",">"); }
		return o.join("");
	}
	function make_html_row(ws/*:Worksheet*/, r/*:Range*/, R/*:number*/, o/*:Sheet2HTMLOpts*/)/*:string*/ {
		var M/*:Array<Range>*/ = (ws['!merges'] ||[]);
		var oo/*:Array<string>*/ = [];
		for(var C = r.s.c; C <= r.e.c; ++C) {
			var RS = 0, CS = 0;
			for(var j = 0; j < M.length; ++j) {
				if(M[j].s.r > R || M[j].s.c > C) continue;
				if(M[j].e.r < R || M[j].e.c < C) continue;
				if(M[j].s.r < R || M[j].s.c < C) { RS = -1; break; }
				RS = M[j].e.r - M[j].s.r + 1; CS = M[j].e.c - M[j].s.c + 1; break;
			}
			if(RS < 0) continue;
			var coord = encode_cell({r:R,c:C});
			var cell = o.dense ? (ws[R]||[])[C] : ws[coord];
			/* TODO: html entities */
			var w = (cell && cell.v != null) && (cell.h || escapehtml(cell.w || (format_cell(cell), cell.w) || "")) || "";
			var sp = ({}/*:any*/);
			if(RS > 1) sp.rowspan = RS;
			if(CS > 1) sp.colspan = CS;
			if(o.editable) w = '<span contenteditable="true">' + w + '</span>';
			else if(cell) {
				sp["data-t"] = cell && cell.t || 'z';
				if(cell.v != null) sp["data-v"] = cell.v;
				if(cell.z != null) sp["data-z"] = cell.z;
				if(cell.l && (cell.l.Target || "#").charAt(0) != "#") w = '<a href="' + cell.l.Target +'">' + w + '</a>';
			}
			sp.id = (o.id || "sjs") + "-" + coord;
			oo.push(writextag('td', w, sp));
		}
		var preamble = "<tr>";
		return preamble + oo.join("") + "</tr>";
	}

	var HTML_BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
	var HTML_END = '</body></html>';

	function make_html_preamble(ws/*:Worksheet*/, R/*:Range*/, o/*:Sheet2HTMLOpts*/)/*:string*/ {
		var out/*:Array<string>*/ = [];
		return out.join("") + '<table' + (o && o.id ? ' id="' + o.id + '"' : "") + '>';
	}

	function sheet_to_html(ws/*:Worksheet*/, opts/*:?Sheet2HTMLOpts*//*, wb:?Workbook*/)/*:string*/ {
		var o = opts || {};
		var header = o.header != null ? o.header : HTML_BEGIN;
		var footer = o.footer != null ? o.footer : HTML_END;
		var out/*:Array<string>*/ = [header];
		var r = decode_range(ws['!ref']);
		o.dense = Array.isArray(ws);
		out.push(make_html_preamble(ws, r, o));
		for(var R = r.s.r; R <= r.e.r; ++R) out.push(make_html_row(ws, r, R, o));
		out.push("</table>" + footer);
		return out.join("");
	}

	function sheet_add_dom(ws/*:Worksheet*/, table/*:HTMLElement*/, _opts/*:?any*/)/*:Worksheet*/ {
		var opts = _opts || {};
		var or_R = 0, or_C = 0;
		if(opts.origin != null) {
			if(typeof opts.origin == 'number') or_R = opts.origin;
			else {
				var _origin/*:CellAddress*/ = typeof opts.origin == "string" ? decode_cell(opts.origin) : opts.origin;
				or_R = _origin.r; or_C = _origin.c;
			}
		}

		var rows/*:HTMLCollection<HTMLTableRowElement>*/ = table.getElementsByTagName('tr');
		var sheetRows = Math.min(opts.sheetRows||10000000, rows.length);
		var range/*:Range*/ = {s:{r:0,c:0},e:{r:or_R,c:or_C}};
		if(ws["!ref"]) {
			var _range/*:Range*/ = decode_range(ws["!ref"]);
			range.s.r = Math.min(range.s.r, _range.s.r);
			range.s.c = Math.min(range.s.c, _range.s.c);
			range.e.r = Math.max(range.e.r, _range.e.r);
			range.e.c = Math.max(range.e.c, _range.e.c);
			if(or_R == -1) range.e.r = or_R = _range.e.r + 1;
		}
		var merges/*:Array<Range>*/ = [], midx = 0;
		var rowinfo/*:Array<RowInfo>*/ = ws["!rows"] || (ws["!rows"] = []);
		var _R = 0, R = 0, _C = 0, C = 0, RS = 0, CS = 0;
		if(!ws["!cols"]) ws['!cols'] = [];
		for(; _R < rows.length && R < sheetRows; ++_R) {
			var row/*:HTMLTableRowElement*/ = rows[_R];
			if (is_dom_element_hidden(row)) {
				if (opts.display) continue;
				rowinfo[R] = {hidden: true};
			}
			var elts/*:HTMLCollection<HTMLTableCellElement>*/ = (row.children/*:any*/);
			for(_C = C = 0; _C < elts.length; ++_C) {
				var elt/*:HTMLTableCellElement*/ = elts[_C];
				if (opts.display && is_dom_element_hidden(elt)) continue;
				var v/*:?string*/ = elt.hasAttribute('data-v') ? elt.getAttribute('data-v') : elt.hasAttribute('v') ? elt.getAttribute('v') : htmldecode(elt.innerHTML);
				var z/*:?string*/ = elt.getAttribute('data-z') || elt.getAttribute('z');
				for(midx = 0; midx < merges.length; ++midx) {
					var m/*:Range*/ = merges[midx];
					if(m.s.c == C + or_C && m.s.r < R + or_R && R + or_R <= m.e.r) { C = m.e.c+1 - or_C; midx = -1; }
				}
				/* TODO: figure out how to extract nonstandard mso- style */
				CS = +elt.getAttribute("colspan") || 1;
				if( ((RS = (+elt.getAttribute("rowspan") || 1)))>1 || CS>1) merges.push({s:{r:R + or_R,c:C + or_C},e:{r:R + or_R + (RS||1) - 1, c:C + or_C + (CS||1) - 1}});
				var o/*:Cell*/ = {t:'s', v:v};
				var _t/*:string*/ = elt.getAttribute("data-t") || elt.getAttribute("t") || "";
				if(v != null) {
					if(v.length == 0) o.t = _t || 'z';
					else if(opts.raw || v.trim().length == 0 || _t == "s");
					else if(v === 'TRUE') o = {t:'b', v:true};
					else if(v === 'FALSE') o = {t:'b', v:false};
					else if(!isNaN(fuzzynum(v))) o = {t:'n', v:fuzzynum(v)};
					else if(!isNaN(fuzzydate(v).getDate())) {
						o = ({t:'d', v:parseDate(v)}/*:any*/);
						if(!opts.cellDates) o = ({t:'n', v:datenum(o.v)}/*:any*/);
						o.z = opts.dateNF || table_fmt[14];
					}
				}
				if(o.z === undefined && z != null) o.z = z;
				/* The first link is used.  Links are assumed to be fully specified.
				 * TODO: The right way to process relative links is to make a new <a> */
				var l = "", Aelts = elt.getElementsByTagName("A");
				if(Aelts && Aelts.length) for(var Aelti = 0; Aelti < Aelts.length; ++Aelti)	if(Aelts[Aelti].hasAttribute("href")) {
					l = Aelts[Aelti].getAttribute("href"); if(l.charAt(0) != "#") break;
				}
				if(l && l.charAt(0) != "#") o.l = ({ Target: l });
				if(opts.dense) { if(!ws[R + or_R]) ws[R + or_R] = []; ws[R + or_R][C + or_C] = o; }
				else ws[encode_cell({c:C + or_C, r:R + or_R})] = o;
				if(range.e.c < C + or_C) range.e.c = C + or_C;
				C += CS;
			}
			++R;
		}
		if(merges.length) ws['!merges'] = (ws["!merges"] || []).concat(merges);
		range.e.r = Math.max(range.e.r, R - 1 + or_R);
		ws['!ref'] = encode_range(range);
		if(R >= sheetRows) ws['!fullref'] = encode_range((range.e.r = rows.length-_R+R-1 + or_R,range)); // We can count the real number of rows to parse but we don't to improve the performance
		return ws;
	}

	function parse_dom_table(table/*:HTMLElement*/, _opts/*:?any*/)/*:Worksheet*/ {
		var opts = _opts || {};
		var ws/*:Worksheet*/ = opts.dense ? ([]/*:any*/) : ({}/*:any*/);
		return sheet_add_dom(ws, table, _opts);
	}

	function table_to_book(table/*:HTMLElement*/, opts/*:?any*/)/*:Workbook*/ {
		return sheet_to_workbook(parse_dom_table(table, opts), opts);
	}

	function is_dom_element_hidden(element/*:HTMLElement*/)/*:boolean*/ {
		var display/*:string*/ = '';
		var get_computed_style/*:?function*/ = get_get_computed_style_function(element);
		if(get_computed_style) display = get_computed_style(element).getPropertyValue('display');
		if(!display) display = element.style && element.style.display;
		return display === 'none';
	}

	/* global getComputedStyle */
	function get_get_computed_style_function(element/*:HTMLElement*/)/*:?function*/ {
		// The proper getComputedStyle implementation is the one defined in the element window
		if(element.ownerDocument.defaultView && typeof element.ownerDocument.defaultView.getComputedStyle === 'function') return element.ownerDocument.defaultView.getComputedStyle;
		// If it is not available, try to get one from the global namespace
		if(typeof getComputedStyle === 'function') return getComputedStyle;
		return null;
	}
	function fix_opts_func(defaults/*:Array<Array<any> >*/)/*:{(o:any):void}*/ {
		return function fix_opts(opts) {
			for(var i = 0; i != defaults.length; ++i) {
				var d = defaults[i];
				if(opts[d[0]] === undefined) opts[d[0]] = d[1];
				if(d[2] === 'n') opts[d[0]] = Number(opts[d[0]]);
			}
		};
	}

	function fix_write_opts(opts) {
	fix_opts_func([
		['cellDates', false], /* write date cells with type `d` */

		['bookSST', false], /* Generate Shared String Table */

		['bookType', 'xlsx'], /* Type of workbook (xlsx/m/b) */

		['compression', false], /* Use file compression */

		['WTF', false] /* WTF mode (throws errors) */
	])(opts);
	}

	function write_zip_xlsx(wb/*:Workbook*/, opts/*:WriteOpts*/)/*:ZIP*/ {
		_shapeid = 1024;
		if(wb && !wb.SSF) {
			wb.SSF = dup(table_fmt);
		}
		if(wb && wb.SSF) {
			make_ssf(); SSF_load_table(wb.SSF);
			// $FlowIgnore
			opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
			opts.ssf = wb.SSF;
		}
		opts.rels = {}; opts.wbrels = {};
		opts.Strings = /*::((*/[]/*:: :any):SST)*/; opts.Strings.Count = 0; opts.Strings.Unique = 0;
		if(browser_has_Map) opts.revStrings = new Map();
		else { opts.revStrings = {}; opts.revStrings.foo = []; delete opts.revStrings.foo; }
		var wbext = "xml";
		var vbafmt = VBAFMTS.indexOf(opts.bookType) > -1;
		var ct = new_ct();
		fix_write_opts(opts = opts || {});
		var zip = zip_new();
		var f = "", rId = 0;

		opts.cellXfs = [];
		get_cell_style(opts.cellXfs, {}, {revssf:{"General":0}});

		if(!wb.Props) wb.Props = {};

		f = "docProps/core.xml";
		zip_add_file(zip, f, write_core_props(wb.Props, opts));
		ct.coreprops.push(f);
		add_rels(opts.rels, 2, f, RELS.CORE_PROPS);

		/*::if(!wb.Props) throw "unreachable"; */
		f = "docProps/app.xml";
		if(wb.Props && wb.Props.SheetNames);
		else if(!wb.Workbook || !wb.Workbook.Sheets) wb.Props.SheetNames = wb.SheetNames;
		else {
			var _sn = [];
			for(var _i = 0; _i < wb.SheetNames.length; ++_i)
				if((wb.Workbook.Sheets[_i]||{}).Hidden != 2) _sn.push(wb.SheetNames[_i]);
			wb.Props.SheetNames = _sn;
		}
		wb.Props.Worksheets = wb.Props.SheetNames.length;
		zip_add_file(zip, f, write_ext_props(wb.Props));
		ct.extprops.push(f);
		add_rels(opts.rels, 3, f, RELS.EXT_PROPS);

		if(wb.Custprops !== wb.Props && keys(wb.Custprops||{}).length > 0) {
			f = "docProps/custom.xml";
			zip_add_file(zip, f, write_cust_props(wb.Custprops));
			ct.custprops.push(f);
			add_rels(opts.rels, 4, f, RELS.CUST_PROPS);
		}

		var people = ["SheetJ5"];
		opts.tcid = 0;

		for(rId=1;rId <= wb.SheetNames.length; ++rId) {
			var wsrels = {'!id':{}};
			var ws = wb.Sheets[wb.SheetNames[rId-1]];
			var _type = (ws || {})["!type"] || "sheet";
			switch(_type) {
			case "chart":
				/* falls through */
			default:
				f = "xl/worksheets/sheet" + rId + "." + wbext;
				zip_add_file(zip, f, write_ws_xml(rId-1, opts, wb, wsrels));
				ct.sheets.push(f);
				add_rels(opts.wbrels, -1, "worksheets/sheet" + rId + "." + wbext, RELS.WS[0]);
			}

			if(ws) {
				var comments = ws['!comments'];
				var need_vml = false;
				var cf = "";
				if(comments && comments.length > 0) {
					var needtc = false;
					comments.forEach(function(carr) {
						carr[1].forEach(function(c) { if(c.T == true) needtc = true; });
					});
					if(needtc) {
						cf = "xl/threadedComments/threadedComment" + rId + "." + wbext;
						zip_add_file(zip, cf, write_tcmnt_xml(comments, people, opts));
						ct.threadedcomments.push(cf);
						add_rels(wsrels, -1, "../threadedComments/threadedComment" + rId + "." + wbext, RELS.TCMNT);
					}

					cf = "xl/comments" + rId + "." + wbext;
					zip_add_file(zip, cf, write_comments_xml(comments));
					ct.comments.push(cf);
					add_rels(wsrels, -1, "../comments" + rId + "." + wbext, RELS.CMNT);
					need_vml = true;
				}
				if(ws['!legacy']) {
					if(need_vml) zip_add_file(zip, "xl/drawings/vmlDrawing" + (rId) + ".vml", write_comments_vml(rId, ws['!comments']));
				}
				delete ws['!comments'];
				delete ws['!legacy'];
			}

			if(wsrels['!id'].rId1) zip_add_file(zip, get_rels_path(f), write_rels(wsrels));
		}

		if(opts.Strings != null && opts.Strings.length > 0) {
			f = "xl/sharedStrings." + wbext;
			zip_add_file(zip, f, write_sst_xml(opts.Strings, opts));
			ct.strs.push(f);
			add_rels(opts.wbrels, -1, "sharedStrings." + wbext, RELS.SST);
		}

		f = "xl/workbook." + wbext;
		zip_add_file(zip, f, write_wb_xml(wb));
		ct.workbooks.push(f);
		add_rels(opts.rels, 1, f, RELS.WB);

		/* TODO: something more intelligent with themes */

		f = "xl/theme/theme1.xml";
		zip_add_file(zip, f, write_theme(wb.Themes, opts));
		ct.themes.push(f);
		add_rels(opts.wbrels, -1, "theme/theme1.xml", RELS.THEME);

		/* TODO: something more intelligent with styles */

		f = "xl/styles." + wbext;
		zip_add_file(zip, f, write_sty_xml(wb, opts));
		ct.styles.push(f);
		add_rels(opts.wbrels, -1, "styles." + wbext, RELS.STY);

		if(wb.vbaraw && vbafmt) {
			f = "xl/vbaProject.bin";
			zip_add_file(zip, f, wb.vbaraw);
			ct.vba.push(f);
			add_rels(opts.wbrels, -1, "vbaProject.bin", RELS.VBA);
		}

		f = "xl/metadata." + wbext;
		zip_add_file(zip, f, write_xlmeta_xml());
		ct.metadata.push(f);
		add_rels(opts.wbrels, -1, "metadata." + wbext, RELS.XLMETA);

		if(people.length > 1) {
			f = "xl/persons/person.xml";
			zip_add_file(zip, f, write_people_xml(people));
			ct.people.push(f);
			add_rels(opts.wbrels, -1, "persons/person.xml", RELS.PEOPLE);
		}

		zip_add_file(zip, "[Content_Types].xml", write_ct(ct, opts));
		zip_add_file(zip, '_rels/.rels', write_rels(opts.rels));
		zip_add_file(zip, 'xl/_rels/workbook.' + wbext + '.rels', write_rels(opts.wbrels));

		delete opts.revssf; delete opts.ssf;
		return zip;
	}
	function write_cfb_ctr(cfb/*:CFBContainer*/, o/*:WriteOpts*/)/*:any*/ {
		switch(o.type) {
			case "base64": case "binary": break;
			case "buffer": case "array": o.type = ""; break;
			case "file": return write_dl(o.file, CFB.write(cfb, {type:has_buf ? 'buffer' : ""}));
			case "string": throw new Error("'string' output type invalid for '" + o.bookType + "' files");
			default: throw new Error("Unrecognized type " + o.type);
		}
		return CFB.write(cfb, o);
	}
	function write_zip_typeXLSX(wb/*:Workbook*/, opts/*:?WriteOpts*/)/*:any*/ {
		var o = dup(opts||{});
		var z = write_zip_xlsx(wb, o);
		return write_zip_denouement(z, o);
	}
	function write_zip_denouement(z/*:any*/, o/*:?WriteOpts*/)/*:any*/ {
		var oopts = {};
		var ftype = has_buf ? "nodebuffer" : (typeof Uint8Array !== "undefined" ? "array" : "string");
		if(o.compression) oopts.compression = 'DEFLATE';
		if(o.password) oopts.type = ftype;
		else switch(o.type) {
			case "base64": oopts.type = "base64"; break;
			case "binary": oopts.type = "string"; break;
			case "string": throw new Error("'string' output type invalid for '" + o.bookType + "' files");
			case "buffer":
			case "file": oopts.type = ftype; break;
			default: throw new Error("Unrecognized type " + o.type);
		}
		var out = z.FullPaths ? CFB.write(z, {fileType:"zip", type: /*::(*/{"nodebuffer": "buffer", "string": "binary"}/*:: :any)*/[oopts.type] || oopts.type, compression: !!o.compression}) : z.generate(oopts);
		if(typeof Deno !== "undefined") {
			if(typeof out == "string") {
				if(o.type == "binary" || o.type == "base64") return out;
				out = new Uint8Array(s2ab(out));
			}
		}
	/*jshint -W083 */
		if(o.password && typeof encrypt_agile !== 'undefined') return write_cfb_ctr(encrypt_agile(out, o.password), o); // eslint-disable-line no-undef
	/*jshint +W083 */
		if(o.type === "file") return write_dl(o.file, out);
		return o.type == "string" ? utf8read(/*::(*/out/*:: :any)*/) : out;
	}

	function writeSyncXLSX(wb/*:Workbook*/, opts/*:?WriteOpts*/) {
		check_wb(wb);
		var o = dup(opts||{});
		if(o.cellStyles) { o.cellNF = true; o.sheetStubs = true; }
		if(o.type == "array") { o.type = "binary"; var out/*:string*/ = (writeSyncXLSX(wb, o)/*:any*/); o.type = "array"; return s2ab(out); }
		return write_zip_typeXLSX(wb, o);
	}

	function resolve_book_type(o/*:WriteFileOpts*/) {
		if(o.bookType) return;
		var _BT = {
			"xls": "biff8",
			"htm": "html",
			"slk": "sylk",
			"socialcalc": "eth",
			"Sh33tJS": "WTF"
		};
		var ext = o.file.slice(o.file.lastIndexOf(".")).toLowerCase();
		if(ext.match(/^\.[a-z]+$/)) o.bookType = ext.slice(1);
		o.bookType = _BT[o.bookType] || o.bookType;
	}

	function writeFileSyncXLSX(wb/*:Workbook*/, filename/*:string*/, opts/*:?WriteFileOpts*/) {
		var o = opts||{}; o.type = 'file';
		o.file = filename;
		resolve_book_type(o);
		return writeSyncXLSX(wb, o);
	}
	/*::
	type MJRObject = {
		row: any;
		isempty: boolean;
	};
	*/
	function make_json_row(sheet/*:Worksheet*/, r/*:Range*/, R/*:number*/, cols/*:Array<string>*/, header/*:number*/, hdr/*:Array<any>*/, dense/*:boolean*/, o/*:Sheet2JSONOpts*/)/*:MJRObject*/ {
		var rr = encode_row(R);
		var defval = o.defval, raw = o.raw || !Object.prototype.hasOwnProperty.call(o, "raw");
		var isempty = true;
		var row/*:any*/ = (header === 1) ? [] : {};
		if(header !== 1) {
			if(Object.defineProperty) try { Object.defineProperty(row, '__rowNum__', {value:R, enumerable:false}); } catch(e) { row.__rowNum__ = R; }
			else row.__rowNum__ = R;
		}
		if(!dense || sheet[R]) for (var C = r.s.c; C <= r.e.c; ++C) {
			var val = dense ? sheet[R][C] : sheet[cols[C] + rr];
			if(val === undefined || val.t === undefined) {
				if(defval === undefined) continue;
				if(hdr[C] != null) { row[hdr[C]] = defval; }
				continue;
			}
			var v = val.v;
			switch(val.t){
				case 'z': if(v == null) break; continue;
				case 'e': v = (v == 0 ? null : void 0); break;
				case 's': case 'd': case 'b': case 'n': break;
				default: throw new Error('unrecognized type ' + val.t);
			}
			if(hdr[C] != null) {
				if(v == null) {
					if(val.t == "e" && v === null) row[hdr[C]] = null;
					else if(defval !== undefined) row[hdr[C]] = defval;
					else if(raw && v === null) row[hdr[C]] = null;
					else continue;
				} else {
					row[hdr[C]] = raw && (val.t !== "n" || (val.t === "n" && o.rawNumbers !== false)) ? v : format_cell(val,v,o);
				}
				if(v != null) isempty = false;
			}
		}
		return { row: row, isempty: isempty };
	}


	function sheet_to_json(sheet/*:Worksheet*/, opts/*:?Sheet2JSONOpts*/) {
		if(sheet == null || sheet["!ref"] == null) return [];
		var val = {t:'n',v:0}, header = 0, offset = 1, hdr/*:Array<any>*/ = [], v=0, vv="";
		var r = {s:{r:0,c:0},e:{r:0,c:0}};
		var o = opts || {};
		var range = o.range != null ? o.range : sheet["!ref"];
		if(o.header === 1) header = 1;
		else if(o.header === "A") header = 2;
		else if(Array.isArray(o.header)) header = 3;
		else if(o.header == null) header = 0;
		switch(typeof range) {
			case 'string': r = safe_decode_range(range); break;
			case 'number': r = safe_decode_range(sheet["!ref"]); r.s.r = range; break;
			default: r = range;
		}
		if(header > 0) offset = 0;
		var rr = encode_row(r.s.r);
		var cols/*:Array<string>*/ = [];
		var out/*:Array<any>*/ = [];
		var outi = 0, counter = 0;
		var dense = Array.isArray(sheet);
		var R = r.s.r, C = 0;
		var header_cnt = {};
		if(dense && !sheet[R]) sheet[R] = [];
		var colinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!cols"] || [];
		var rowinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!rows"] || [];
		for(C = r.s.c; C <= r.e.c; ++C) {
			if(((colinfo[C]||{}).hidden)) continue;
			cols[C] = encode_col(C);
			val = dense ? sheet[R][C] : sheet[cols[C] + rr];
			switch(header) {
				case 1: hdr[C] = C - r.s.c; break;
				case 2: hdr[C] = cols[C]; break;
				case 3: hdr[C] = o.header[C - r.s.c]; break;
				default:
					if(val == null) val = {w: "__EMPTY", t: "s"};
					vv = v = format_cell(val, null, o);
					counter = header_cnt[v] || 0;
					if(!counter) header_cnt[v] = 1;
					else {
						do { vv = v + "_" + (counter++); } while(header_cnt[vv]); header_cnt[v] = counter;
						header_cnt[vv] = 1;
					}
					hdr[C] = vv;
			}
		}
		for (R = r.s.r + offset; R <= r.e.r; ++R) {
			if ((rowinfo[R]||{}).hidden) continue;
			var row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
			if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) out[outi++] = row.row;
		}
		out.length = outi;
		return out;
	}

	var qreg = /"/g;
	function make_csv_row(sheet/*:Worksheet*/, r/*:Range*/, R/*:number*/, cols/*:Array<string>*/, fs/*:number*/, rs/*:number*/, FS/*:string*/, o/*:Sheet2CSVOpts*/)/*:?string*/ {
		var isempty = true;
		var row/*:Array<string>*/ = [], txt = "", rr = encode_row(R);
		for(var C = r.s.c; C <= r.e.c; ++C) {
			if (!cols[C]) continue;
			var val = o.dense ? (sheet[R]||[])[C]: sheet[cols[C] + rr];
			if(val == null) txt = "";
			else if(val.v != null) {
				isempty = false;
				txt = ''+(o.rawNumbers && val.t == "n" ? val.v : format_cell(val, null, o));
				for(var i = 0, cc = 0; i !== txt.length; ++i) if((cc = txt.charCodeAt(i)) === fs || cc === rs || cc === 34 || o.forceQuotes) {txt = "\"" + txt.replace(qreg, '""') + "\""; break; }
				if(txt == "ID") txt = '"ID"';
			} else if(val.f != null && !val.F) {
				isempty = false;
				txt = '=' + val.f; if(txt.indexOf(",") >= 0) txt = '"' + txt.replace(qreg, '""') + '"';
			} else txt = "";
			/* NOTE: Excel CSV does not support array formulae */
			row.push(txt);
		}
		if(o.blankrows === false && isempty) return null;
		return row.join(FS);
	}

	function sheet_to_csv(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/)/*:string*/ {
		var out/*:Array<string>*/ = [];
		var o = opts == null ? {} : opts;
		if(sheet == null || sheet["!ref"] == null) return "";
		var r = safe_decode_range(sheet["!ref"]);
		var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
		var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
		var endregex = new RegExp((FS=="|" ? "\\|" : FS)+"+$");
		var row = "", cols/*:Array<string>*/ = [];
		o.dense = Array.isArray(sheet);
		var colinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!cols"] || [];
		var rowinfo/*:Array<ColInfo>*/ = o.skipHidden && sheet["!rows"] || [];
		for(var C = r.s.c; C <= r.e.c; ++C) if (!((colinfo[C]||{}).hidden)) cols[C] = encode_col(C);
		var w = 0;
		for(var R = r.s.r; R <= r.e.r; ++R) {
			if ((rowinfo[R]||{}).hidden) continue;
			row = make_csv_row(sheet, r, R, cols, fs, rs, FS, o);
			if(row == null) { continue; }
			if(o.strip) row = row.replace(endregex,"");
			if(row || (o.blankrows !== false)) out.push((w++ ? RS : "") + row);
		}
		delete o.dense;
		return out.join("");
	}

	function sheet_to_txt(sheet/*:Worksheet*/, opts/*:?Sheet2CSVOpts*/) {
		if(!opts) opts = {}; opts.FS = "\t"; opts.RS = "\n";
		var s = sheet_to_csv(sheet, opts);
		return s;
	}

	function sheet_to_formulae(sheet/*:Worksheet*/)/*:Array<string>*/ {
		var y = "", x, val="";
		if(sheet == null || sheet["!ref"] == null) return [];
		var r = safe_decode_range(sheet['!ref']), rr = "", cols/*:Array<string>*/ = [], C;
		var cmds/*:Array<string>*/ = [];
		var dense = Array.isArray(sheet);
		for(C = r.s.c; C <= r.e.c; ++C) cols[C] = encode_col(C);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			rr = encode_row(R);
			for(C = r.s.c; C <= r.e.c; ++C) {
				y = cols[C] + rr;
				x = dense ? (sheet[R]||[])[C] : sheet[y];
				val = "";
				if(x === undefined) continue;
				else if(x.F != null) {
					y = x.F;
					if(!x.f) continue;
					val = x.f;
					if(y.indexOf(":") == -1) y = y + ":" + y;
				}
				if(x.f != null) val = x.f;
				else if(x.t == 'z') continue;
				else if(x.t == 'n' && x.v != null) val = "" + x.v;
				else if(x.t == 'b') val = x.v ? "TRUE" : "FALSE";
				else if(x.w !== undefined) val = "'" + x.w;
				else if(x.v === undefined) continue;
				else if(x.t == 's') val = "'" + x.v;
				else val = ""+x.v;
				cmds[cmds.length] = y + "=" + val;
			}
		}
		return cmds;
	}

	function sheet_add_json(_ws/*:?Worksheet*/, js/*:Array<any>*/, opts)/*:Worksheet*/ {
		var o = opts || {};
		var offset = +!o.skipHeader;
		var ws/*:Worksheet*/ = _ws || ({}/*:any*/);
		var _R = 0, _C = 0;
		if(ws && o.origin != null) {
			if(typeof o.origin == 'number') _R = o.origin;
			else {
				var _origin/*:CellAddress*/ = typeof o.origin == "string" ? decode_cell(o.origin) : o.origin;
				_R = _origin.r; _C = _origin.c;
			}
		}
		var cell/*:Cell*/;
		var range/*:Range*/ = ({s: {c:0, r:0}, e: {c:_C, r:_R + js.length - 1 + offset}}/*:any*/);
		if(ws['!ref']) {
			var _range = safe_decode_range(ws['!ref']);
			range.e.c = Math.max(range.e.c, _range.e.c);
			range.e.r = Math.max(range.e.r, _range.e.r);
			if(_R == -1) { _R = _range.e.r + 1; range.e.r = _R + js.length - 1 + offset; }
		} else {
			if(_R == -1) { _R = 0; range.e.r = js.length - 1 + offset; }
		}
		var hdr/*:Array<string>*/ = o.header || [], C = 0;

		js.forEach(function (JS, R/*:number*/) {
			keys(JS).forEach(function(k) {
				if((C=hdr.indexOf(k)) == -1) hdr[C=hdr.length] = k;
				var v = JS[k];
				var t = 'z';
				var z = "";
				var ref = encode_cell({c:_C + C,r:_R + R + offset});
				cell = ws_get_cell_stub(ws, ref);
				if(v && typeof v === 'object' && !(v instanceof Date)){
					ws[ref] = v;
				} else {
					if(typeof v == 'number') t = 'n';
					else if(typeof v == 'boolean') t = 'b';
					else if(typeof v == 'string') t = 's';
					else if(v instanceof Date) {
						t = 'd';
						if(!o.cellDates) { t = 'n'; v = datenum(v); }
						z = (o.dateNF || table_fmt[14]);
					}
					else if(v === null && o.nullError) { t = 'e'; v = 0; }
					if(!cell) ws[ref] = cell = ({t:t, v:v}/*:any*/);
					else {
						cell.t = t; cell.v = v;
						delete cell.w; delete cell.R;
						if(z) cell.z = z;
					}
					if(z) cell.z = z;
				}
			});
		});
		range.e.c = Math.max(range.e.c, _C + hdr.length - 1);
		var __R = encode_row(_R);
		if(offset) for(C = 0; C < hdr.length; ++C) ws[encode_col(C + _C) + __R] = {t:'s', v:hdr[C]};
		ws['!ref'] = encode_range(range);
		return ws;
	}
	function json_to_sheet(js/*:Array<any>*/, opts)/*:Worksheet*/ { return sheet_add_json(null, js, opts); }

	/* get cell, creating a stub if necessary */
	function ws_get_cell_stub(ws/*:Worksheet*/, R, C/*:?number*/)/*:Cell*/ {
		/* A1 cell address */
		if(typeof R == "string") {
			/* dense */
			if(Array.isArray(ws)) {
				var RC = decode_cell(R);
				if(!ws[RC.r]) ws[RC.r] = [];
				return ws[RC.r][RC.c] || (ws[RC.r][RC.c] = {t:'z'});
			}
			return ws[R] || (ws[R] = {t:'z'});
		}
		/* cell address object */
		if(typeof R != "number") return ws_get_cell_stub(ws, encode_cell(R));
		/* R and C are 0-based indices */
		return ws_get_cell_stub(ws, encode_cell({r:R,c:C||0}));
	}

	/* find sheet index for given name / validate index */
	function wb_sheet_idx(wb/*:Workbook*/, sh/*:number|string*/) {
		if(typeof sh == "number") {
			if(sh >= 0 && wb.SheetNames.length > sh) return sh;
			throw new Error("Cannot find sheet # " + sh);
		} else if(typeof sh == "string") {
			var idx = wb.SheetNames.indexOf(sh);
			if(idx > -1) return idx;
			throw new Error("Cannot find sheet name |" + sh + "|");
		} else throw new Error("Cannot find sheet |" + sh + "|");
	}

	/* simple blank workbook object */
	function book_new()/*:Workbook*/ {
		return { SheetNames: [], Sheets: {} };
	}

	/* add a worksheet to the end of a given workbook */
	function book_append_sheet(wb/*:Workbook*/, ws/*:Worksheet*/, name/*:?string*/, roll/*:?boolean*/)/*:string*/ {
		var i = 1;
		if(!name) for(; i <= 0xFFFF; ++i, name = undefined) if(wb.SheetNames.indexOf(name = "Sheet" + i) == -1) break;
		if(!name || wb.SheetNames.length >= 0xFFFF) throw new Error("Too many worksheets");
		if(roll && wb.SheetNames.indexOf(name) >= 0) {
			var m = name.match(/(^.*?)(\d+)$/);
			i = m && +m[2] || 0;
			var root = m && m[1] || name;
			for(++i; i <= 0xFFFF; ++i) if(wb.SheetNames.indexOf(name = root + i) == -1) break;
		}
		check_ws_name(name);
		if(wb.SheetNames.indexOf(name) >= 0) throw new Error("Worksheet with name |" + name + "| already exists!");

		wb.SheetNames.push(name);
		wb.Sheets[name] = ws;
		return name;
	}

	/* set sheet visibility (visible/hidden/very hidden) */
	function book_set_sheet_visibility(wb/*:Workbook*/, sh/*:number|string*/, vis/*:number*/) {
		if(!wb.Workbook) wb.Workbook = {};
		if(!wb.Workbook.Sheets) wb.Workbook.Sheets = [];

		var idx = wb_sheet_idx(wb, sh);
		// $FlowIgnore
		if(!wb.Workbook.Sheets[idx]) wb.Workbook.Sheets[idx] = {};

		switch(vis) {
			case 0: case 1: case 2: break;
			default: throw new Error("Bad sheet visibility setting " + vis);
		}
		// $FlowIgnore
		wb.Workbook.Sheets[idx].Hidden = vis;
	}

	/* set number format */
	function cell_set_number_format(cell/*:Cell*/, fmt/*:string|number*/) {
		cell.z = fmt;
		return cell;
	}

	/* set cell hyperlink */
	function cell_set_hyperlink(cell/*:Cell*/, target/*:string*/, tooltip/*:?string*/) {
		if(!target) {
			delete cell.l;
		} else {
			cell.l = ({ Target: target }/*:Hyperlink*/);
			if(tooltip) cell.l.Tooltip = tooltip;
		}
		return cell;
	}
	function cell_set_internal_link(cell/*:Cell*/, range/*:string*/, tooltip/*:?string*/) { return cell_set_hyperlink(cell, "#" + range, tooltip); }

	/* add to cell comments */
	function cell_add_comment(cell/*:Cell*/, text/*:string*/, author/*:?string*/) {
		if(!cell.c) cell.c = [];
		cell.c.push({t:text, a:author||"SheetJS"});
	}

	/* set array formula and flush related cells */
	function sheet_set_array_formula(ws/*:Worksheet*/, range, formula/*:string*/, dynamic/*:boolean*/) {
		var rng = typeof range != "string" ? range : safe_decode_range(range);
		var rngstr = typeof range == "string" ? range : encode_range(range);
		for(var R = rng.s.r; R <= rng.e.r; ++R) for(var C = rng.s.c; C <= rng.e.c; ++C) {
			var cell = ws_get_cell_stub(ws, R, C);
			cell.t = 'n';
			cell.F = rngstr;
			delete cell.v;
			if(R == rng.s.r && C == rng.s.c) {
				cell.f = formula;
				if(dynamic) cell.D = true;
			}
		}
		return ws;
	}

	var utils/*:any*/ = {
		encode_col: encode_col,
		encode_row: encode_row,
		encode_cell: encode_cell,
		encode_range: encode_range,
		decode_col: decode_col,
		decode_row: decode_row,
		split_cell: split_cell,
		decode_cell: decode_cell,
		decode_range: decode_range,
		format_cell: format_cell,
		sheet_add_aoa: sheet_add_aoa,
		sheet_add_json: sheet_add_json,
		sheet_add_dom: sheet_add_dom,
		aoa_to_sheet: aoa_to_sheet,
		json_to_sheet: json_to_sheet,
		table_to_sheet: parse_dom_table,
		table_to_book: table_to_book,
		sheet_to_csv: sheet_to_csv,
		sheet_to_txt: sheet_to_txt,
		sheet_to_json: sheet_to_json,
		sheet_to_html: sheet_to_html,
		sheet_to_formulae: sheet_to_formulae,
		sheet_to_row_object_array: sheet_to_json,
		sheet_get_cell: ws_get_cell_stub,
		book_new: book_new,
		book_append_sheet: book_append_sheet,
		book_set_sheet_visibility: book_set_sheet_visibility,
		cell_set_number_format: cell_set_number_format,
		cell_set_hyperlink: cell_set_hyperlink,
		cell_set_internal_link: cell_set_internal_link,
		cell_add_comment: cell_add_comment,
		sheet_set_array_formula: sheet_set_array_formula,
		consts: {
			SHEET_VISIBLE: 0,
			SHEET_HIDDEN: 1,
			SHEET_VERY_HIDDEN: 2
		}
	};

	var Direction;
	(function (Direction) {
	  Direction[Direction["Row"] = 0] = "Row";
	  Direction[Direction["Column"] = 1] = "Column";
	})(Direction || (Direction = {}));
	const json2excel = props => {
	  // 工作薄
	  const workbook = utils.book_new();
	  props.data.forEach(item => {
	    // 工作表
	    const worksheet = utils.json_to_sheet([]);
	    // 头部标题
	    item.sheetData.forEach((subItem, subIndex) => {
	      utils.sheet_add_aoa(worksheet, Object.entries(subItem), {
	        origin: `A${subIndex + 1}`
	      });
	    });
	    // 增加工作表
	    if (item.sheetName.length >= 31) {
	      utils.book_append_sheet(workbook, worksheet, `${item.sheetName.substr(0, 28)}...`);
	    } else {
	      utils.book_append_sheet(workbook, worksheet, item.sheetName);
	    }
	    // 计算列宽度
	    const cellMaxWidth = item.sheetData.reduce((maxWidth, cell) => {
	      const cellEntries = Object.entries(cell);
	      return {
	        colWidth1: Math.max(cellEntries[0][0].length, maxWidth.colWidth1),
	        colWidth2: Math.max(cellEntries[0][1].length, maxWidth.colWidth2)
	      };
	    }, {
	      colWidth1: 10,
	      colWidth2: 10
	    });
	    worksheet['!cols'] = [{
	      wch: cellMaxWidth.colWidth1
	    }, {
	      wch: cellMaxWidth.colWidth2
	    }];
	  });
	  // 写入到文件
	  writeFileSyncXLSX(workbook, props.fileName);
	};

	return json2excel;

}));
