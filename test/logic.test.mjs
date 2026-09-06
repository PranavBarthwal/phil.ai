import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import assert from 'node:assert/strict';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const src = readFileSync(join(root, 'apps-script/Code.gs'), 'utf8');

function extractFunction(name) {
  const start = src.search(new RegExp('function\\s+' + name + '\\s*\\('));
  if (start < 0) throw new Error('missing function ' + name);
  let i = src.indexOf('{', start);
  let depth = 0;
  for (; i < src.length; i++) {
    if (src[i] === '{') depth++;
    else if (src[i] === '}') {
      depth--;
      if (depth === 0) return src.slice(start, i + 1);
    }
  }
  throw new Error('unclosed function ' + name);
}

const fns = [
  'parseTemplateMarker',
  'parseAddressList',
  'joinAddresses',
  'personalize',
  'parseGroqContent',
  'textToHtml',
  'escapeHtml',
  'htmlToPlain',
  'pickSubject',
  'chosenTemplateId',
  'normalizeStatus'
];
const sandbox = { CONFIG: { DEFAULT_TEMPLATE_ID: 'default', FOLLOWUP_TEMPLATE_IDS: ['followup_1', 'followup_2'] } };
const body = fns.map(extractFunction).join('\n');
const runner = new Function('CONFIG', body + '\nreturn { ' + fns.join(', ') + ' };');
const lib = runner(sandbox.CONFIG);

assert.equal(lib.parseTemplateMarker('Template: default'), 'default');
assert.equal(lib.parseTemplateMarker('=== TEMPLATE: XYZ_Experience ==='), 'xyz_experience');
assert.equal(lib.parseTemplateMarker('Subject: hello'), null);

assert.deepEqual(lib.parseAddressList('a@x.com; b@y.com, not-an-email'), ['a@x.com', 'b@y.com']);
assert.equal(lib.joinAddresses(['a@x.com', 'b@y.com']), 'a@x.com, b@y.com');

assert.equal(lib.personalize('Hi {{first_name}} at {{company}}', { first_name: 'Ada', company: 'Acme' }), 'Hi Ada at Acme');
assert.equal(lib.personalize('keep {{unknown}}', {}), 'keep {{unknown}}');

const groq = lib.parseGroqContent('```json\n{"subjects":["A","B"],"body_html":"<p>Hi</p>"}\n```');
assert.deepEqual(groq.subjects, ['A', 'B']);
assert.equal(groq.bodyHtml, '<p>Hi</p>');

const row = ['', '2'];
const col = { subjectChoice: 1 };
assert.equal(lib.pickSubject(['one', 'two', 'three'], row, col), 'two');
assert.equal(lib.pickSubject(['one', 'two'], [''], { subjectChoice: 0 }), 'one');

const templates = { default: {}, followup_1: {}, cool: {} };
assert.equal(lib.chosenTemplateId(['cool'], { template: 0 }, templates), 'cool');
assert.equal(lib.chosenTemplateId([''], { template: 0 }, { followup_1: {}, cool: {} }), 'cool');

assert.equal(lib.normalizeStatus(' Sent '), 'sent');
assert.ok(lib.htmlToPlain('<p>Hi &amp; bye</p>').indexOf('Hi & bye') !== -1);

console.log('ok');
