import muze from "@viz/muze";
import "@viz/muze/muze.css";
import { exportToExcel } from "../index.js";

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

let { schema, data } = await loadData({
  dataSetLink: "/data/superstore/data.json",
  schemaLink: "/data/superstore/schema.json",
});

const { DataModel } = muze;
const env = muze();

const formattedData = DataModel.loadDataSync(data, schema);
let rootData = new DataModel(formattedData);

window.canvas = env
  .canvas()
  .data(rootData)
  .width(900)
  .height(600)
  .rows(["Sub-Category", "Ship Mode"])
  .columns(["Ship Mode"])
  .layers([
    {
      mark: "text",
      encoding: {
        text: "Quantity",
        color: "Region",
      },
    },
  ])
  .mount("#chart");

const button = document.getElementById("button");
button.addEventListener("click", (event) => {
  exportToExcel(window.canvas);
});
