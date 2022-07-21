#!/usr/bin/env python3
from docx import Document
from urlextract import URLExtract
import argparse


def main():
    parser = argparse.ArgumentParser(description='Extract links from a docx file.')
    parser.add_argument('--file', metavar='FILE', type=str, help='input file', required=True)
    args = parser.parse_args()

    document = Document(args.file)
    extractor = URLExtract()

    urls = set()

    for para in document.paragraphs:
        rs = para._element.xpath('.//w:t')
        paragraph_text_decoded = (u" ".join([r.text for r in rs]))
        for url in extractor.find_urls(paragraph_text_decoded):
            urls.add(url)

    urls = sorted(urls)
    urls = ("\n".join(urls))

    print(urls)


if __name__ == '__main__':
    main()
