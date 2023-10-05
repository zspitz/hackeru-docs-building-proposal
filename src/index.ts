import * as fs from 'fs'
import { Tokens, marked } from 'marked';
import pptxgen from 'pptxgenjs';
import PptxGenJS from 'pptxgenjs';
import { ParentState, defaultParentState, getTableRows, getTextPropObjects } from './tokenToPptxGen.js';

const md = fs.readFileSync('README.md', 'utf-8');

const pres: PptxGenJS.default = new (pptxgen as any)();
pres.rtlMode = true;
let currentSection = '(Default)';
let currentSlide: PptxGenJS.default.Slide;

const tokens = marked.Lexer.lex(md);

const addIntroSlide = (token: Tokens.Heading) => {
    let fontSize = 10;
    if (token.depth === 1) {
        fontSize = 40;
    } else if (token.depth === 2) {
        currentSection = token.text;
        fontSize = 30;
    }
    pres.addSection({ title: currentSection });

    const texts = getTextPropObjects(token)

    const introSlide = pres.addSlide({
        sectionTitle: currentSection || ''
    });
    introSlide.addText(texts, {
        lang: 'he',
        fontSize,
        rtlMode: true,
        w: '100%',
        h: '100%',
        align: 'center',
        valign: 'middle',
        bold: true
    });
}

// make sure there is a current slide
const ensureCurrentSlide = () => {
    if (currentSlide) { return; }
    currentSlide = pres.addSlide({
        sectionTitle: currentSection
    })
}

// dump current texts into current/new slide
const writeTexts = () => {
    if (texts.length === 0) { return; }
    ensureCurrentSlide()
    currentSlide.addText(texts, {
        lang: 'he',
        rtlMode: true,
        x: .1,
        y: .1,
        h: '95%',
        w: '95%',
        valign: 'top',
        align: 'right'
    });
    texts = [];
}

let texts: PptxGenJS.default.TextProps[] = [];

tokens.forEach((token, index) => {
    if (token.type === 'heading' && token.depth <= 3) {
        writeTexts()

        // create intro slide for depth 1 and 2
        if (token.depth === 1 || token.depth == 2) {
            addIntroSlide(token as Tokens.Heading)
        }

        currentSlide = pres.addSlide({
            sectionTitle: currentSection
        })
    }

    const parentState: ParentState =
        token.type === 'heading' ?
            { ...defaultParentState, strongDepth: 1 } :
            defaultParentState

    if (token.type === 'heading' && token.depth === 1) { return; }

    if (token.type === 'table') {
        // append table to current slide
        ensureCurrentSlide()
        currentSlide.addTable(getTableRows(token as Tokens.Table))
    } else {
        getTextPropObjects(token, texts, parentState)
    }
})

// add remaining texts into current/new slide
writeTexts()

// breakline in addText objects - line break before or after?

pres.writeFile({
    fileName: './dist/output.pptx'
})