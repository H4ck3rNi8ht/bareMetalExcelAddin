/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    myfunction();
  }
});

export async function myfunction() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */

    });
  } catch (error) {
    console.error(error);
  }
}
