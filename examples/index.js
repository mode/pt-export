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
  dataSetLink: "/data/mn_mv/new_data.csv",
  schemaLink: "/data/mn_mv/schema.json",
});

const { DataModel } = muze;
const env = muze();
// console.log("Hi");

const formattedData = DataModel.loadDataSync(data, schema);
let rootData = new DataModel(formattedData);

rootData = rootData.select({
  operator: "and",
  conditions: [
    {
      field: "region",
      value: ["Central", "East"],
      operator: "in",
    },
    // {
    //   field: "Measure Names",
    //   value: ["quantity", "profit"],
    //   operator: "in",
    // },
    {
      field: "ship_mode",
      value: ["First Class", "Second Class"],
      operator: "in",
    },
    {
      field: "segment",
      value: ["Consumer", "Corporate"],
      operator: "in",
    },
  ],
});

window.canvas = env
  .canvas()
  .data(rootData)
  .width(900)
  .height(600)
  .rows([["Measure Names", "region", "ship_mode"]])
  .columns([["region", "category"]])
  // .detail(["category"])
  .layers([
    {
      mark: "text",
      encoding: {
        text: "Measure Values",
        color: "sales",
      },
    },
  ])
  .mount("#chart");

const button = document.getElementById("button");
button.addEventListener("click", (event) => {
  exportToExcel(window.canvas);
});
