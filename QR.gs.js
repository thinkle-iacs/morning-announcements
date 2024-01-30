function addQR() {
  let selection = SlidesApp.getActivePresentation().getSelection();
  if (!selection) {
    SlidesApp.getUi().alert('Nothing is selected');
  }
  let elements = selection.getPageElementRange();
  if (elements) {
    for (let e of elements.getPageElements()) {
      let links = e.asShape().getText().getLinks();
      let yoffset = 0;
      for (let l of links) {
        let url = l.getTextStyle().getLink().getUrl();
        let qrurl = `https://api.qrserver.com/v1/create-qr-code/?data=${url}&size=400x400`;
        SlidesApp.getActivePresentation().getSelection().getCurrentPage().insertImage(
          qrurl,0,yoffset,400/links.length,400/links.length
        );
        yoffset += 400/links.length;
      }
      }
  } else {
    SlidesApp.getUi().alert('Did you forget to select some text with a link?');
  }  
}
