import muze from "@viz/muze";
import "@viz/muze/muze.css";
import { exportToExcel } from "../main.js";

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
  // .columns([["Category"],["Sub-Category"]])
  // .rows(["Region" ,"Ship Mode"])
  .rows(["Sub-Category", "Ship Mode"])
  .columns(["Ship Mode"])
  // .color('Region')
  .layers([
    {
      mark: "text",
      encoding: {
        text: "Quantity",
        color: "Region",
        // backgroundColor: "Region"
      },
    },
  ])
  // .layers([
  //   {
  //     mark: "text",
  //     encoding: {
  //       text: "Quantity",
  //       color: "Region",
  //     },
  //   },
  // ])
  // .config({
  //   legend: {
  //     color: {
  //       fields: {
  //         Region: {
  //           domainRangeMap: {
  //             "Central": "#251991",
  //             "West": "#c7db0f",
  //           },
  //           range: [
  //             "#37B067",
  //             "#6296BC",
  //             "#EDB40D",
  //             "#7FD7C1",
  //             "#9F8CAE",
  //             "#EB6672",
  //             "#376C72",
  //             "#EE9DCC",
  //           ],
  //         },
  //       },
  //     },
  //   },
  // })
  .mount("#chart");

const button = document.getElementById("button");
button.addEventListener("click", (event) => {
  exportToExcel(window.canvas);
});
