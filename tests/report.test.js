import test from 'node:test';
import assert from 'node:assert/strict';
import {reportBlocks,reportHTML} from '../public/report.js';
const report={name:'test.xlsx',generatedAt:'2026-10-22T10:00:00Z',version:'2.0',threshold:'1000',sha256:'abc',sheets:[{name:'Other',supported:false,reason:'قالب نامعتبر'}],checks:[],crossChecks:[],summary:{records:0,subtotals:0,confirmed:0,affected:0},findings:[]};
test('report exposes skipped sheet and names user-authored notes',()=>{const text=reportBlocks(report,{inspector:'بازرس',note:'یادداشت انسانی'}).map(x=>x.text).join('\n');assert.ok(text.includes('بررسی نشده'));assert.ok(text.includes('یادداشت بازرس (افزوده کاربر)'));assert.ok(text.includes('یادداشت انسانی'));assert.ok(text.includes('بازرس'));});
test('report escapes workbook and user text',()=>{const html=reportHTML({...report,name:'<img src=x onerror=alert(1)>'},{note:'<script>alert(1)</script>'});assert.ok(!html.includes('<img'));assert.ok(!html.includes('<script'));assert.ok(html.includes('&lt;script&gt;'));});
