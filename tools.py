# The COPYRIGHT file at the top level of this repository contains
# the full copyright notices and license terms.
from html.parser import HTMLParser

import markdown


def safe_text(value):
    return '' if value is None else str(value)


class SafeDict(dict):

    def __missing__(self, key):
        return ''


class TemplateRecord:

    def __init__(self, record):
        self._record = record

    def __getattr__(self, name):
        return getattr(self._record, name)

    def __str__(self):
        return safe_text(getattr(self._record, 'rec_name', self._record))

    def __bool__(self):
        return bool(self._record)


class MarkdownParagraphParser(HTMLParser):

    def __init__(self):
        super().__init__()
        self.blocks = []
        self._current_block = None
        self._current_runs = []
        self._block_style = {}
        self._bold = 0
        self._italic = 0

    def handle_starttag(self, tag, attrs):
        if tag in {'p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'}:
            self._start_block(tag)
        elif tag in {'strong', 'b'}:
            self._bold += 1
        elif tag in {'em', 'i'}:
            self._italic += 1
        elif tag == 'br':
            self._current_runs.append({'break': True})

    def handle_endtag(self, tag):
        if tag in {'p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'}:
            self._finish_block()
        elif tag in {'strong', 'b'} and self._bold:
            self._bold -= 1
        elif tag in {'em', 'i'} and self._italic:
            self._italic -= 1

    def handle_data(self, data):
        if not data:
            return
        if self._current_block is None:
            self._start_block('p')
        self._current_runs.append({
                'text': data,
                'bold': self._block_style.get('bold', False) or self._bold > 0,
                'italic': self._italic > 0,
                })

    def _start_block(self, tag):
        if self._current_block is not None:
            self._finish_block()
        self._current_block = tag
        self._current_runs = []
        self._block_style = {'bold': tag in {'h1', 'h2', 'h3', 'h4', 'h5', 'h6'}}

    def _finish_block(self):
        if self._current_block is None:
            return
        runs = [r for r in self._current_runs
            if r.get('break') or r.get('text')]
        if runs:
            self.blocks.append({
                    'runs': runs,
                    'bullet': self._current_block == 'li',
                    })
        self._current_block = None
        self._current_runs = []
        self._block_style = {}


def markdown_to_paragraphs(text):
    html = markdown.markdown(text or '')
    parser = MarkdownParagraphParser()
    parser.feed(html)
    parser.close()
    return parser.blocks or [{'text': ''}]
