const d = document;
const $main = d.querySelector("main");

fetch("markdowns/00_intro.md")
  .then((res) => (res.ok ? res.text() : Promise.reject(res)))
  .then((text) => {
    $main.innerHTML = new showdown.Converter().makeHtml(text);
  })
  .catch((err) => {
    let message = err.statusText || "Ocurrio un error";
    $main.innerHTML = `Error ${err.status} : ${message}`;
  });
