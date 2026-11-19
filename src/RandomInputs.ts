function main(workbook: ExcelScript.Workbook) {

  const sheet = workbook.getWorksheet("Architect");
  const arrowrange = sheet.getRangeByIndexes(4, 1, 50, 5);
  arrowrange.setValue("");
  const messagerane = sheet.getRangeByIndexes(5, 6, 50, 2);
  messagerane.setValue("");
  sheet.getCell(4, 7).setValue(1);
  sheet.getCell(4, 6).setValue("");

  const startRow = 4;
  const startCol = 1;
  const minRows = 5;
  const maxRows = 20;
  const dimMin = 4;
  const dimMax = 11;
  const batchMin = 3;
  const batchMax = 11;
  const nonLinears = ["ReLU", "Tanh", "GELU", "Softmax", "LayerNorm"];
  const firstChoices = ["Rand", "Rand", "Rand", "Image", "tinyDigits"];

  function randInt(a: number, b: number): number {
    return Math.floor(Math.random() * (b - a + 1)) + a;
  }
  function pick<T>(arr: T[]): T {
    return arr[randInt(0, arr.length - 1)];
  }
  function isNonLinear(layer: string): boolean {
    return nonLinears.includes(layer);
  }

  function pickHiddenLayer(prev: string): string {
    // LayerNorm must be immediately followed by Linear so post-processing can
    // insert it correctly (checks layers[i-1] === "LayerNorm" before a Linear).
    if (prev === "LayerNorm") {
      return "Linear";
    }
    let options = [
      "Linear", "Linear", "Linear", "Linear", "Linear",
      "ReLU", "Tanh", "GELU",
      "LayerNorm"  // Same frequency as ReLU, Tanh, GELU
    ];
    if (prev === "Linear") {
      options = options.filter(x => x !== "Linear" && x !== "LayerNorm");
    }
    if (isNonLinear(prev)) {
      options = options.filter(x => x !== prev);
    }
    return pick(options);
  }

  function pickFinalBeforeLoss(prev: string): string {
    let options = ["Softmax", "ReLU", "Tanh", "GELU"];
    if (isNonLinear(prev)) {
      options = options.filter(x => x !== prev);
    }
    return pick(options);
  }

  const totalRows = randInt(minRows, maxRows);
  let layers: string[] = [];
  const firstLayer = pick(firstChoices);
  layers.push(firstLayer);

  if (firstLayer === "Image" || firstLayer === "tinyDigits") {
    layers.push("Flatten");
  }

  while (layers.length < totalRows - 1) {
    const prev = layers[layers.length - 1];
    const remainingBeforeLoss = (totalRows - 1) - layers.length;
    if (remainingBeforeLoss === 1) {
      layers.push(pickFinalBeforeLoss(prev));
      break;
    }
    layers.push(pickHiddenLayer(prev));
  }

  const loss = pick(["CELoss", "MSELoss"]);
  if (loss === "CELoss") {
    layers[layers.length - 1] = "Softmax";
    if (layers.length >= 2 && layers[layers.length - 2] === "Softmax") {
      layers[layers.length - 2] = pick(["ReLU", "Tanh", "GELU"]);
    }
  } else {
    if (layers[layers.length - 1] === "Linear") {
      const prev = layers.length >= 2 ? layers[layers.length - 2] : "";
      layers[layers.length - 1] = pickFinalBeforeLoss(prev);
    }
  }
  layers.push(loss);

  // Post-process: insert LayerNorm BEFORE Linear,
  // but NOT if the layer after Linear is a Loss layer
  let processedLayers: string[] = [];
  for (let i = 0; i < layers.length; i++) {
    if (layers[i] === "LayerNorm") {
      // Skip placeholder — handled when we encounter the next Linear
      continue;
    }
    if (layers[i] === "Linear") {
      // Check if previous layer in original was LayerNorm
      if (i > 0 && layers[i - 1] === "LayerNorm") {
        // Only insert LayerNorm if the layer AFTER this Linear is NOT a Loss
        const nextLayer = i + 1 < layers.length ? layers[i + 1] : "";
        const isBeforeLoss = nextLayer === "CELoss" || nextLayer === "MSELoss";
        if (!isBeforeLoss) {
          processedLayers.push("LayerNorm");
        }
      }
      processedLayers.push("Linear");
    } else {
      processedLayers.push(layers[i]);
    }
  }
  layers = processedLayers;

  // Recalculate linear indices after post-processing
  const linearIndices: number[] = [];
  for (let i = 0; i < layers.length; i++) {
    if (layers[i] === "Linear") {
      linearIndices.push(i);
    }
  }
  const lastLinearIndex =
    linearIndices.length > 0 ? linearIndices[linearIndices.length - 1] : -1;

  let output: (string | number)[][] = [];
  for (let i = 0; i < layers.length; i++) {
    const layer = layers[i];
    let dim: string | number = "";
    let batch: string | number = "";

    if (layer === "tinyDigits") {
      dim = 5;
    } else if (layer === "Image") {
      dim = randInt(2, 5);
    } else if (layer === "Rand" || layer === "Linear") {
      if (firstLayer === "tinyDigits" && layer === "Linear" && i === lastLinearIndex) {
        dim = 10;
      } else {
        dim = randInt(dimMin, dimMax);
      }
    }

    if (i === 0) {
      batch = randInt(batchMin, batchMax);
    }
    output.push([layer, dim, batch]);
  }

  sheet.getRangeByIndexes(startRow, startCol, maxRows, 3)
    .clear(ExcelScript.ClearApplyTo.contents);
  sheet.getRangeByIndexes(startRow, startCol, output.length, 3)
    .setValues(output);
}

// function main(workbook: ExcelScript.Workbook) {
  
//   const sheet = workbook.getWorksheet("Architect");

//   const arrowrange = sheet.getRangeByIndexes(4, 1, 50, 5);
//   arrowrange.setValue("");

//   const messagerane = sheet.getRangeByIndexes(5, 6, 50, 2);
//   messagerane.setValue("");
//   sheet.getCell(4, 7).setValue(1);
//   sheet.getCell(4, 6).setValue(false);



//   const startRow = 4;
//   const startCol = 1;

//   const minRows = 5;
//   const maxRows = 17;

//   const dimMin = 4;
//   const dimMax = 11;

//   const batchMin = 3;
//   const batchMax = 11;

//   const nonLinears = ["ReLU", "Tanh", "GELU", "Softmax"];
//   const firstChoices = ["Rand", "Rand", "Rand", "Image", "tinyDigits"];

//   function randInt(a: number, b: number): number {
//     return Math.floor(Math.random() * (b - a + 1)) + a;
//   }

//   function pick<T>(arr: T[]): T {
//     return arr[randInt(0, arr.length - 1)];
//   }

//   function isNonLinear(layer: string): boolean {
//     return nonLinears.includes(layer);
//   }

//   function pickHiddenLayer(prev: string): string {
//     let options = [
//       "Linear", "Linear", "Linear", "Linear", "Linear",
//       "ReLU", "Tanh", "GELU"
//     ];

//     if (prev === "Linear") {
//       options = options.filter(x => x !== "Linear");
//     }

//     if (isNonLinear(prev)) {
//       options = options.filter(x => x !== prev);
//     }

//     return pick(options);
//   }

//   function pickFinalBeforeLoss(prev: string): string {
//     let options = ["Softmax", "ReLU", "Tanh", "GELU"];

//     if (isNonLinear(prev)) {
//       options = options.filter(x => x !== prev);
//     }

//     return pick(options);
//   }

//   const totalRows = randInt(minRows, maxRows);
//   let layers: string[] = [];

//   const firstLayer = pick(firstChoices);
//   layers.push(firstLayer);

//   if (firstLayer === "Image" || firstLayer === "tinyDigits") {
//     layers.push("Flatten");
//   }

//   while (layers.length < totalRows - 1) {
//     const prev = layers[layers.length - 1];
//     const remainingBeforeLoss = (totalRows - 1) - layers.length;

//     if (remainingBeforeLoss === 1) {
//       layers.push(pickFinalBeforeLoss(prev));
//       break;
//     }

//     layers.push(pickHiddenLayer(prev));
//   }

//   const loss = pick(["CELoss", "MSELoss"]);

//   if (loss === "CELoss") {
//     layers[layers.length - 1] = "Softmax";

//     if (layers.length >= 2 && layers[layers.length - 2] === "Softmax") {
//       layers[layers.length - 2] = pick(["ReLU", "Tanh", "GELU"]);
//     }
//   } else {
//     if (layers[layers.length - 1] === "Linear") {
//       const prev = layers.length >= 2 ? layers[layers.length - 2] : "";
//       layers[layers.length - 1] = pickFinalBeforeLoss(prev);
//     }
//   }

//   layers.push(loss);

//   const linearIndices: number[] = [];
//   for (let i = 0; i < layers.length; i++) {
//     if (layers[i] === "Linear") {
//       linearIndices.push(i);
//     }
//   }

//   const lastLinearIndex =
//     linearIndices.length > 0 ? linearIndices[linearIndices.length - 1] : -1;

//   let output: (string | number)[][] = [];

//   for (let i = 0; i < layers.length; i++) {
//     const layer = layers[i];
//     let dim: string | number = "";
//     let batch: string | number = "";

//     if (layer === "tinyDigits") {
//       dim = 5;
//     } else if (layer === "Image") {
//       dim = randInt(2, 5);
//     } else if (layer === "Rand" || layer === "Linear") {
//       if (firstLayer === "tinyDigits" && layer === "Linear" && i === lastLinearIndex) {
//         dim = 10;
//       } else {
//         dim = randInt(dimMin, dimMax);
//       }
//     }

//     if (i === 0) {
//       batch = randInt(batchMin, batchMax);
//     }

//     output.push([layer, dim, batch]);
//   }

//   sheet.getRangeByIndexes(startRow, startCol, maxRows, 3)
//     .clear(ExcelScript.ClearApplyTo.contents);

//   sheet.getRangeByIndexes(startRow, startCol, output.length, 3)
//     .setValues(output);
// }