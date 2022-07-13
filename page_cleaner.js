// remove everything after the content
while (document.body.lastElementChild && (
    document.body.lastElementChild.tagName != 'DIV' ||
    document.body.lastElementChild.id != 'body'
)) {
   document.body.removeChild(document.body.lastElementChild);
}
// remove everything before the content
while (document.body.firstElementChild && (
    document.body.firstElementChild.tagName != 'DIV' ||
    document.body.firstElementChild.id != 'body'
)) {
    document.body.removeChild(document.body.firstElementChild);
}