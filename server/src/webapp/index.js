function getQueryParameters(queryString) {
  const regex = new RegExp("[\\?&]([a-zA-Z0-9_-]+)=([^&#]*)", "g");
  let match = null;
  let result = {};
  do {
    match = regex.exec(queryString);
    if (match !== null) {
      let name = match[1];
      let value = decodeURIComponent(match[2].replace(/\+/g, " "));
      result[name] = value;
    }
  } while (match !== null);
  return result;
}

function removeAllChildren(node) {
  while (node.hasChildNodes()) {
    node.removeChild(node.lastChild);
  }
}

function loadVersion() {
  fetch("/version")
    .then((response) => {
      if (response.status !== 200) {
        response.text().then((text) => console.error(text));
      } else {
        return response.json();
      }
    })
    .then((versionData) => {
      document
        .getElementById("version")
        .appendChild(
          document.createTextNode(
            `${
              versionData.tag.length === 0 ? "latest" : versionData.tag
            }, git ${versionData.commit}`
          )
        );
    });
}
