import muze from "@viz/muze";
import "@viz/muze/muze.css";
import { exportToExcel  } from "../main";
import { server } from "@vitest/browser/context";

const { readFile, writeFile, removeFile } = server.commands;

const loadData = async function ({
  dataSetLink = "/data/cars.json",
  schemaLink = "/data/cars-schema.json",
}) {
  let data = await fetch(dataSetLink).then((d) =>
    dataSetLink.split(".").pop() === "csv" ? d.text() : d.json()
  );
  let schema = await fetch(schemaLink).then((d) => d.json());

  return { schema, data };
};

export default async function HelloWorld({ name }) {
  const parent = document.createElement("div");
  document.body.appendChild(parent);

  const button = document.createElement("button");
  button.textContent = 'Click here';

  document.body.appendChild(button);

  button.addEventListener("click", (event) => {
  exportToExcel(window.canvas);
});

  const dataContent = await readFile(
    "../examples/public/data/superstore/data.json"
  );
  const data = JSON.parse(dataContent);
  
  const schemaContent = await readFile(
    "../examples/public/data/superstore/schema.json"
  );
  const schema = JSON.parse(schemaContent);

  const { DataModel } = muze;
  const env = muze();

  const formattedData = DataModel.loadDataSync(data, schema);
  let rootData = new DataModel(formattedData);

  window.canvas = env
    .canvas()
    .data(rootData)
    .width(900)
    .height(600)
    // .columns([["Category"],["Sub-Category"]])
    // .rows(["Region" ,"Ship Mode"])
    .rows(["Sub-Category", "Ship Mode"])
    .columns([["Ship Mode"], ["Category"]])
    // .color('Region')
    .layers([
      {
        mark: "text",
        encoding: {
          text: "Quantity",
          backgroundColor: "Region",
        },
      },
    ])
    .mount(parent);

  // canvas.once("afterRendered", () => {
  //   console.log("rendered");
  // });
  // const h1 = document.createElement('h1')
  // h1.textContent = 'Hello ' + name + '!'
  // parent.appendChild(h1)

  return parent;
}
