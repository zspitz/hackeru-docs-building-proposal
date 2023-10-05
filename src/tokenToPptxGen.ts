import { Token, Tokens, marked } from 'marked';
import PptxGenJS from 'pptxgenjs';

type DepthName = 'strongDepth' | 'delDepth' | 'emDepth'
export interface ParentState {
    readonly delDepth: number,
    readonly emDepth: number,
    readonly strongDepth: number,
    readonly listDepth: number
}

type HasSubtokens = Token & { tokens: Token[] };
const isHasSubtokens = (token: Token): token is HasSubtokens =>
    (token as HasSubtokens).tokens !== undefined;

type HasText = Token & { text: string };
const isHasText = (token: Token): token is HasText =>
    (token as HasText).text !== undefined;

export const defaultParentState: ParentState = {
    delDepth: 0,
    emDepth: 0,
    strongDepth: 0,
    listDepth: -1
}

export const getTextPropObjects = (
    token: Token,
    texts: PptxGenJS.default.TextProps[] = [],
    parentState: ParentState = defaultParentState
) => {

    switch (token.type) {
        // ignore space token
        case 'space':
            return texts

        // handle list
        case 'list':
            // list can be ordered or unordered
            // ordered list can start at a specific number
            const bulletProps: true | Exclude<PptxGenJS.default.TextProps['options'], undefined>['bullet'] =
                token.ordered ?
                    {
                        type: 'number',
                        numberStartAt: token.start
                    } :
                    true
            // per the documentation, it should be possible to pass an object for unordered lists as well
            // but the bullets don't show unless explicitly using true

            parentState = {
                ...parentState,
                listDepth: parentState.listDepth + 1
            }

            token.items.forEach((listitem: Tokens.ListItem) => {
                const listitemTexts = getTextPropObjects(listitem, undefined, parentState)
                listitemTexts.forEach((listitemText, index) => {
                    listitemText.options ??= {}
                    if (index === 0) {
                        listitemText.options.bullet = bulletProps
                        listitemText.options.indentLevel = parentState.listDepth
                    }

                    // if last text object, set breakline to true; otherwise set to false
                    listitemText.options.breakLine =
                        index === listitemTexts.length - 1
                })
                listitemTexts.forEach(x => texts.push(x))
            })
            return texts;

        // tokens that cannot be parsed to text prop objects
        case 'table':
            throw 'Cannot parse Markdown table into text prop objects. Use getTableRows instead.'
        case 'image':
            throw 'Cannot parse Markdown image into text prop objects. Use getImageProps instead.'
        case 'hr':
        case 'def':
            throw `Token of type '${token.type}' cannot be parsed to text prop objects`
    }

    // if current node's state is different from parent state (e.g. strong or del)
    // create copy of parent state and modify
    switch (token.type) {
        case 'strong':
        case 'del':
        case 'em':
            const key: DepthName = `${token.type}Depth`
            parentState = {
                ...parentState,
                [key]: parentState[key] + 1
            }
            break;
    }

    // if has subtokens
    // Block elements: Heading, Blockquote, ListItem, Paragraph
    // Inlines: TableCell, Text with subtokens, Link, Strong, Em, Del
    if (isHasSubtokens(token)) {

        // walk each subtoken with new parent state
        token.tokens.forEach(subtoken => getTextPropObjects(subtoken, texts, parentState))
        switch (token.type) {
            case 'heading':
            case 'blockquote':
            case 'listitem':
            case 'paragraph':
                const last = texts.at(-1)
                if (last) {
                    last.options!.breakLine = true
                }
                break;
        }

    } else if (isHasText(token)) {

        // append new textpropobject based on parent state and current text
        let textPropObject: PptxGenJS.default.TextProps;
        if (token.type === 'text' || token.type === 'codespan') {
            textPropObject = {
                text: token.raw,
                options: {
                    rtlMode: token.type === 'text',
                    bold: parentState.strongDepth > 0,
                    strike: parentState.delDepth > 0 ? 'sngStrike' : undefined,
                    italic: parentState.emDepth > 0
                }
            }
        } else if (token.type === 'code') {
            textPropObject = {
                text: token.text,
                options: {
                    rtlMode: false,
                    fontFace: 'Courier New'
                }
            }
        } else if (token.type === 'html') {
            if (token.text !== '<br/>') { throw 'Only br is allowed as HTML' }
            textPropObject = {
                text: '\n'
            }
        } else {
            throw `Unhandled single-text token of type '${token.type}'`;
        }
        texts.push(textPropObject);

    } else {
        throw `Unhandled token of type '${token.type}'`;
    }

    // certain tokens need to modify the last generated textpropobject -- e.g. line break

    return texts;


    // other: Table, List, HTML, 
    // irrelevant: Hr, Image
    // ?? Def, Escape, Tag, Br, 

    // TODO nested blockquotes
    // TODO nested list
    // TODO number start at

    // walk the token tree; at each node change, create a new object to add to the array


    return [] as PptxGenJS.default.TextProps[];
}

export const getTableRows = (table: Tokens.Table) => {
    const markedRows: Tokens.TableCell[][] = [
        table.header,
        ...table.rows
    ]
    return markedRows.map(row =>
        row.map(cell => {
            const texts: PptxGenJS.default.TextProps[] = [];
            getTextPropObjects(cell as Token, texts)
            return {
                text: texts as any
            }
        })
    )
}

export const getImageProps = (image: Tokens.Image) => {
    throw 'Not implemented'

    return {} as PptxGenJS.default.ImageProps;
}

// export const getPptxGenObjects = (token: Token) => {
//     if (token.type === 'table') { return getTableRows(token as Tokens.Table); }
//     if (token.type === 'image') { return getImageProps(token as Tokens.Image); }

//     // if token has tokens, return TextPropObjects[]
//     switch (token.type) {
//         case 'heading':
//         case 'blockquote':
//         case 'list_item':
//         case 'paragraph':
//         case 'link':
//         case 'strong':
//         case 'em':
//         case 'del':
//             return getTextPropObjects(token as Exclude<typeof token, Tokens.Generic>)

//         case 'text':
//             // TODO what happens if the text token doesn't have a tokens property?
//             return getTextPropObjects(token as Exclude<Exclude<typeof token, Tokens.Generic>, Tokens.Tag>)

//         default:
//         // return single TextPropObject

//     }
// }

// TODO strip breakline from last text before slide