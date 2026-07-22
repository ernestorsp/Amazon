/***************************************************************
 * ASIGNAR RUTAS Y VANS POR SIZE
 *
 * NO consulta la hoja Record.
 * NO depende de vans previamente colocadas en a&aprueba.
 *
 * Lee las rutas AAXI desde STG y las vans OPERACIONALES desde
 * VAN_INFO. Cada van se usa una sola vez y se asigna únicamente
 * a una ruta compatible con su categoría.
 *
 * Los primeros nombres repetidos se revisan en la columna A
 * de la hoja DRIVER.
 *
 * EJEMPLOS:
 * - Si solo existe un John en DRIVER:
 *   John Kent -> JOHN
 *
 * - Si existen varios Clarence en DRIVER:
 *   Clarence Wendell Sneed  -> CLARENCE S
 *   Clarence Gregory Taylor -> CLARENCE T
 *
 * MAPEO DE RUTAS:
 * - Standard Parcel - Custom Delivery Van 16ft        -> CDV
 * - Standard Parcel - Extra Large Van - US           -> XL
 * - Standard Parcel Electric - Rivian MEDIUM         -> EDV
 * - Standard Parcel - Large Van                      -> L
 * - Standard Parcel Step Van - US                    -> SV
 * - Cualquier size que contenga "Nursery"            -> EDV
 *
 * VAN_INFO:
 * - Columna A -> se copia a a&aprueba columna G
 * - Columna B -> número/nombre de van; se copia a columna D
 * - Columna C -> se copia a a&aprueba columna H
 * - Columna D -> size/tipo de van
 * - Columna E -> status; solamente se usa Operational
 *
 * a&aprueba, desde la fila 6:
 * - A: Route
 * - B: Driver
 * - C: Size asignado
 * - D: Van asignada
 * - E: Wave time
 * - F: Line
 * - G: dato VAN_INFO columna A
 * - H: dato VAN_INFO columna C
 * - I: Staging area
 *
 * Este script solamente limpia y escribe A6:I100.
 * No modifica las columnas Q:AF.
 * No modifica filas después de la 100.
 * No agrega filas ni columnas.
 ***************************************************************/

const VAN_ASSIGN_CFG = {
  SHEET_STG: "STG",
  SHEET_VAN_INFO: "VAN_INFO",
  SHEET_OUTPUT: "a&aprueba",
  SHEET_DRIVER: "DRIVER",

  COMPANY_VALUE: "AAXI",
  OUTPUT_START_ROW: 6,
  OUTPUT_END_ROW: 100,

  // STG, columnas 1-based
  STG_COL_COMPANY: 1, // A
  STG_COL_ROUTE: 2,   // B
  STG_COL_SIZE: 4,    // D
  STG_COL_TIME: 5,    // E
  STG_COL_AREA: 6,    // F
  STG_COL_DRIVER: 8,  // H

  // VAN_INFO, columnas 1-based
  VAN_COL_EXTRA_G: 1, // A -> a&aprueba G
  VAN_COL_NUMBER: 2,  // B -> a&aprueba D
  VAN_COL_EXTRA_H: 3, // C -> a&aprueba H
  VAN_COL_SIZE: 4,    // D
  VAN_COL_STATUS: 5,  // E

  // a&aprueba, columnas 1-based
  OUT_COL_ROUTE: 1,   // A
  OUT_COL_DRIVER: 2,  // B
  OUT_COL_SIZE: 3,    // C
  OUT_COL_VAN: 4,     // D
  OUT_COL_TIME: 5,    // E
  OUT_COL_LINE: 6,    // F
  OUT_COL_EXTRA_G: 7, // G
  OUT_COL_EXTRA_H: 8, // H
  OUT_COL_AREA: 9     // I
};

const LINE_BY_TIME = {
  "11:20": "6",
  "11:25": "2",
  "11:30": "2",
  "11:35": "3",
  "11:40": "6",
  "11:50": "5",
  "11:55": "1",
  "12:00": "6",
  "12:15": "6",
  "12:35": "1",
  "12:40": "5"
};

/**
 * FUNCIÓN PRINCIPAL
 */
function llenar_aaprueba_asignando_vans_por_size() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shSTG = ss.getSheetByName(
    VAN_ASSIGN_CFG.SHEET_STG
  );

  const shVan = ss.getSheetByName(
    VAN_ASSIGN_CFG.SHEET_VAN_INFO
  );

  const shOut = ss.getSheetByName(
    VAN_ASSIGN_CFG.SHEET_OUTPUT
  );

  const shDriver = ss.getSheetByName(
    VAN_ASSIGN_CFG.SHEET_DRIVER
  );

  if (!shSTG) {
    throw new Error(
      `No existe la hoja ${VAN_ASSIGN_CFG.SHEET_STG}.`
    );
  }

  if (!shVan) {
    throw new Error(
      `No existe la hoja ${VAN_ASSIGN_CFG.SHEET_VAN_INFO}.`
    );
  }

  if (!shOut) {
    throw new Error(
      `No existe la hoja ${VAN_ASSIGN_CFG.SHEET_OUTPUT}.`
    );
  }

  if (!shDriver) {
    throw new Error(
      `No existe la hoja ${VAN_ASSIGN_CFG.SHEET_DRIVER}.`
    );
  }

  const routes = leerRutasSTG_(shSTG);

  if (routes.length === 0) {
    SpreadsheetApp.getUi().alert(
      `No se encontraron rutas de ` +
      `${VAN_ASSIGN_CFG.COMPANY_VALUE} en STG.`
    );

    return;
  }

  const vansBySize =
    leerVansOperacionalesPorSize_(shVan);

  /*
   * Ordena las rutas primero por la hora del wave
   * y después por su orden original en STG.
   */
  routes.sort((a, b) => {
    const timeDiff =
      parseTimeToMinutes_(a.time) -
      parseTimeToMinutes_(b.time);

    return timeDiff !== 0
      ? timeDiff
      : a.sourceRow - b.sourceRow;
  });

  /*
   * Revisa los primeros nombres repetidos
   * usando la columna A de DRIVER.
   */
  const firstNameCounts =
    contarPrimerosNombresDriver_(shDriver);

  const assignments = [];
  const errors = [];

  routes.forEach(route => {
    route.driverOutput =
      crearNombreDriverSalida_(
        route.driverFull,
        firstNameCounts
      );

    if (!route.requiredSize) {
      errors.push(
        `Ruta ${route.routeNumber || route.routeRaw}: ` +
        `size no reconocido: "${route.sizeText}".`
      );

      assignments.push(
        crearAsignacionSinVan_(route)
      );

      return;
    }

    /*
     * Primero busca una van del size exacto.
     *
     * Si no existe:
     * 1. Intenta usar una EDV.
     * 2. Si tampoco hay EDV, intenta usar una XL.
     * 3. Si tampoco hay XL, intenta usar una L.
     *
     * Cada van continúa utilizándose una sola vez
     * porque se elimina de su pool con shift().
     */
    const vanResult = obtenerVanCompatible_(
      vansBySize,
      route.requiredSize
    );

    const van = vanResult.van;

    if (!van) {
      errors.push(
        `Ruta ${route.routeNumber || route.routeRaw} ` +
        `(${route.driverOutput}): ` +
        `no hay van Operational disponible de size ` +
        `${route.requiredSize}, EDV, XL ni L.`
      );

      assignments.push(
        crearAsignacionSinVan_(route)
      );

      return;
    }

    /*
     * Si se utilizó una van diferente al size original,
     * se registra como advertencia.
     */
    if (vanResult.usedFallback) {
      errors.push(
        `Ruta ${route.routeNumber || route.routeRaw} ` +
        `(${route.driverOutput}): requería ` +
        `${route.requiredSize}, pero se asignó ` +
        `${van.size}.`
      );
    }

    assignments.push({
      routeNumber: route.routeNumber,
      driverOutput: route.driverOutput,

      /*
       * Escribe en columna C el size real
       * de la van asignada.
       */
      assignedSize: van.size,

      vanNumber: van.vanNumber,
      time: route.time,
      line:
        LINE_BY_TIME[toHHMM_(route.time)] || "",
      extraG: van.extraG,
      extraH: van.extraH,
      area: route.area
    });
  });

  escribirAsignaciones_(
    shOut,
    assignments
  );

  const assignedCount = assignments.filter(
    assignment => assignment.vanNumber
  ).length;

  const missingCount =
    assignments.length - assignedCount;

  let message =
    `✅ Asignación terminada.\n\n` +
    `Rutas procesadas: ${assignments.length}\n` +
    `Vans asignadas: ${assignedCount}\n` +
    `Rutas sin van: ${missingCount}\n\n` +
    `Los nombres repetidos fueron revisados ` +
    `en la columna A de DRIVER.\n` +
    `La hoja Record no fue consultada.\n` +
    `Las columnas Q:AF no fueron modificadas.\n` +
    `Las filas después de la 100 no fueron modificadas.`;

  if (errors.length > 0) {
    message +=
      `\n\n⚠️ Revisar:\n- ` +
      `${errors.join("\n- ")}`;
  }

  SpreadsheetApp.getUi().alert(message);
}

/**
 * Lee las rutas de la compañía AAXI desde STG.
 */
function leerRutasSTG_(shSTG) {
  const lastRow = shSTG.getLastRow();

  const lastCol = Math.max(
    shSTG.getLastColumn(),
    VAN_ASSIGN_CFG.STG_COL_DRIVER
  );

  if (lastRow < 2) {
    return [];
  }

  const data = shSTG
    .getRange(
      2,
      1,
      lastRow - 1,
      lastCol
    )
    .getDisplayValues();

  const routes = [];

  data.forEach((row, index) => {
    const company = cleanText_(
      row[
        VAN_ASSIGN_CFG.STG_COL_COMPANY - 1
      ]
    );

    if (
      company.toUpperCase() !==
      VAN_ASSIGN_CFG.COMPANY_VALUE.toUpperCase()
    ) {
      return;
    }

    const routeRaw = cleanText_(
      row[
        VAN_ASSIGN_CFG.STG_COL_ROUTE - 1
      ]
    );

    const sizeText = cleanText_(
      row[
        VAN_ASSIGN_CFG.STG_COL_SIZE - 1
      ]
    );

    const time = cleanText_(
      row[
        VAN_ASSIGN_CFG.STG_COL_TIME - 1
      ]
    );

    const areaRaw = cleanText_(
      row[
        VAN_ASSIGN_CFG.STG_COL_AREA - 1
      ]
    );

    const driverFull = cleanText_(
      row[
        VAN_ASSIGN_CFG.STG_COL_DRIVER - 1
      ]
    );

    /*
     * Si falta ruta, driver u hora,
     * no procesa esa fila.
     */
    if (
      !routeRaw ||
      !driverFull ||
      !time
    ) {
      return;
    }

    routes.push({
      sourceRow: index + 2,
      routeRaw: routeRaw,
      routeNumber:
        extractRouteNumber_(routeRaw),
      sizeText: sizeText,
      requiredSize:
        routeSizeToCode_(sizeText),
      time: time,
      area:
        stripSTGPrefix_(areaRaw),
      driverFull: driverFull
    });
  });

  return routes;
}

/**
 * Lee VAN_INFO y crea grupos de vans Operational
 * por categoría.
 *
 * Cada van solamente puede aparecer una vez.
 */
function leerVansOperacionalesPorSize_(shVan) {
  const lastRow = shVan.getLastRow();

  const pools = crearPoolsVacios_();

  if (lastRow < 2) {
    return pools;
  }

  const data = shVan
    .getRange(
      2,
      1,
      lastRow - 1,
      VAN_ASSIGN_CFG.VAN_COL_STATUS
    )
    .getDisplayValues();

  const seenVans = new Set();

  data.forEach((row, index) => {
    const status = normalizeKey_(
      row[
        VAN_ASSIGN_CFG.VAN_COL_STATUS - 1
      ]
    );

    /*
     * Solo utiliza vans con estado Operational.
     */
    if (status !== "operational") {
      return;
    }

    const vanNumber = cleanText_(
      row[
        VAN_ASSIGN_CFG.VAN_COL_NUMBER - 1
      ]
    );

    const vanSizeText = cleanText_(
      row[
        VAN_ASSIGN_CFG.VAN_COL_SIZE - 1
      ]
    );

    const vanSize =
      vanInfoSizeToCode_(vanSizeText);

    /*
     * Ignora filas sin número de van
     * o sin un size reconocido.
     */
    if (!vanNumber || !vanSize) {
      return;
    }

    const vanKey =
      normalizeKey_(vanNumber);

    /*
     * Detiene el script si encuentra
     * la misma van repetida en VAN_INFO.
     */
    if (seenVans.has(vanKey)) {
      throw new Error(
        `La van "${vanNumber}" está repetida ` +
        `en VAN_INFO. Revisa la fila ` +
        `${index + 2}.`
      );
    }

    seenVans.add(vanKey);

    pools.get(vanSize).push({
      sourceRow: index + 2,
      vanNumber: vanNumber,
      size: vanSize,
      originalSizeText: vanSizeText,

      /*
       * VAN_INFO columna A
       * pasa a a&aprueba columna G.
       */
      extraG:
        row[
          VAN_ASSIGN_CFG.VAN_COL_EXTRA_G - 1
        ],

      /*
       * VAN_INFO columna C
       * pasa a a&aprueba columna H.
       */
      extraH:
        row[
          VAN_ASSIGN_CFG.VAN_COL_EXTRA_H - 1
        ]
    });
  });

  /*
   * Mantiene el orden en que las vans
   * aparecen en VAN_INFO.
   */
  pools.forEach(pool => {
    pool.sort(
      (a, b) =>
        a.sourceRow - b.sourceRow
    );
  });

  return pools;
}

/**
 * Crea los grupos de vans disponibles.
 */
function crearPoolsVacios_() {
  return new Map([
    ["CDV", []],
    ["XL", []],
    ["EDV", []],
    ["L", []],
    ["SV", []]
  ]);
}

/**
 * Obtiene una van compatible.
 *
 * Prioridad:
 * 1. Size exacto requerido.
 * 2. EDV.
 * 3. XL.
 * 4. L.
 *
 * No intenta dos veces el mismo size.
 */
function obtenerVanCompatible_(
  vansBySize,
  requiredSize
) {
  const priority = [
    requiredSize,
    "EDV",
    "XL",
    "L"
  ];

  const checkedSizes = new Set();

  for (const size of priority) {
    if (!size || checkedSizes.has(size)) {
      continue;
    }

    checkedSizes.add(size);

    const pool =
      vansBySize.get(size) || [];

    if (pool.length > 0) {
      const van = pool.shift();

      return {
        van: van,
        usedFallback:
          size !== requiredSize
      };
    }
  }

  return {
    van: null,
    usedFallback: false
  };
}

/**
 * Convierte el size de la ruta de STG
 * a la categoría requerida.
 *
 * Nursery tiene prioridad y siempre devuelve EDV.
 */
function routeSizeToCode_(sizeText) {
  const s =
    normalizeSizeText_(sizeText);

  if (!s) {
    return "";
  }

  /*
   * Cualquier texto que tenga Nursery usa EDV.
   */
  if (s.includes("NURSERY")) {
    return "EDV";
  }

  /*
   * Standard Parcel Electric - Rivian MEDIUM.
   */
  if (
    s.includes(
      "STANDARD PARCEL ELECTRIC"
    ) &&
    s.includes("RIVIAN") &&
    s.includes("MEDIUM")
  ) {
    return "EDV";
  }

  /*
   * Standard Parcel - Custom Delivery Van 16ft.
   */
  if (
    s.includes(
      "CUSTOM DELIVERY VAN"
    ) &&
    s.includes("16FT")
  ) {
    return "CDV";
  }

  /*
   * Standard Parcel - Extra Large Van - US.
   */
  if (
    s.includes("EXTRA LARGE VAN")
  ) {
    return "XL";
  }

  /*
   * Standard Parcel Step Van - US.
   */
  if (
    s.includes("STEP VAN")
  ) {
    return "SV";
  }

  /*
   * Standard Parcel - Large Van.
   *
   * Esta condición debe estar después de
   * EXTRA LARGE VAN para impedir que XL
   * se detecte como L.
   */
  if (
    s.includes("LARGE VAN")
  ) {
    return "L";
  }

  /*
   * También acepta los códigos directamente.
   */
  if (
    [
      "CDV",
      "XL",
      "EDV",
      "L",
      "SV"
    ].includes(s)
  ) {
    return s;
  }

  return "";
}

/**
 * Convierte VAN_INFO columna D a:
 *
 * CDV
 * XL
 * EDV
 * L
 * SV
 *
 * Acepta códigos y descripciones completas.
 */
function vanInfoSizeToCode_(sizeText) {
  const s = normalizeSizeText_(sizeText);

  if (!s) {
    return "";
  }

  /*
   * Códigos exactos.
   */
  if (s === "CDV") return "CDV";
  if (s === "XL") return "XL";
  if (s === "EDV") return "EDV";
  if (s === "L") return "L";
  if (s === "SV") return "SV";

  /*
   * Descripciones completas.
   */
  if (s.includes("NURSERY")) {
    return "EDV";
  }

  if (
    s.includes("RIVIAN") &&
    s.includes("MEDIUM")
  ) {
    return "EDV";
  }

  if (
    s.includes("CUSTOM DELIVERY VAN") &&
    s.includes("16FT")
  ) {
    return "CDV";
  }

  if (s.includes("EXTRA LARGE VAN")) {
    return "XL";
  }

  if (s.includes("STEP VAN")) {
    return "SV";
  }

  /*
   * Debe ir después de EXTRA LARGE VAN.
   */
  if (s.includes("LARGE VAN")) {
    return "L";
  }

  return "";
}

/**
 * Escribe las asignaciones en a&aprueba.
 *
 * Solamente utiliza:
 * - Filas 6 a 100
 * - Columnas A:I
 *
 * No modifica:
 * - Filas después de la 100
 * - Columnas J:AF
 *
 * No inserta filas ni columnas.
 */
function escribirAsignaciones_(shOut, assignments) {
  const startRow =
    VAN_ASSIGN_CFG.OUTPUT_START_ROW;

  const endRow =
    VAN_ASSIGN_CFG.OUTPUT_END_ROW;

  /*
   * 9 columnas = A:I.
   */
  const outputWidth = 9;

  /*
   * Filas disponibles desde la 6 hasta la 100.
   *
   * 100 - 6 + 1 = 95 filas.
   */
  const availableRows =
    endRow - startRow + 1;

  /*
   * Detiene el proceso antes de limpiar o escribir
   * si hay más rutas que filas disponibles.
   *
   * Esto protege las filas después de la 100.
   */
  if (assignments.length > availableRows) {
    throw new Error(
      `Hay ${assignments.length} rutas, pero solamente ` +
      `hay ${availableRows} filas disponibles entre ` +
      `la fila ${startRow} y la fila ${endRow}. ` +
      `No se limpió ni se escribió ningún dato para ` +
      `proteger las fórmulas inferiores.`
    );
  }

  /*
   * Limpia únicamente A6:I100.
   *
   * No utiliza getLastRow(), ya que getLastRow()
   * puede detectar fórmulas ubicadas en otras columnas.
   */
  shOut
    .getRange(
      startRow,
      1,
      availableRows,
      outputWidth
    )
    .clearContent();

  if (assignments.length === 0) {
    return;
  }

  const values = assignments.map(
    assignment => [
      assignment.routeNumber,
      assignment.driverOutput,
      assignment.assignedSize,
      assignment.vanNumber,
      assignment.time,
      assignment.line,
      assignment.extraG,
      assignment.extraH,
      assignment.area
    ]
  );

  /*
   * Escribe solamente la cantidad exacta
   * de rutas procesadas.
   */
  shOut
    .getRange(
      startRow,
      1,
      values.length,
      outputWidth
    )
    .setValues(values);
}

/**
 * Crea una fila aunque no haya van disponible.
 */
function crearAsignacionSinVan_(route) {
  return {
    routeNumber: route.routeNumber,
    driverOutput: route.driverOutput,
    assignedSize: route.requiredSize || "",
    vanNumber: "",
    time: route.time,
    line:
      LINE_BY_TIME[toHHMM_(route.time)] || "",
    extraG: "",
    extraH: "",
    area: route.area
  };
}

/**
 * Cuenta los primeros nombres usando únicamente
 * la columna A de la hoja DRIVER.
 *
 * REGLA:
 *
 * Si JOHN aparece una sola vez:
 * JOHN
 *
 * Si CLARENCE aparece más de una vez:
 * CLARENCE S
 * CLARENCE T
 */
function contarPrimerosNombresDriver_(shDriver) {
  const lastRow = shDriver.getLastRow();
  const counts = new Map();

  if (lastRow < 1) {
    return counts;
  }

  /*
   * Lee toda la columna A de DRIVER.
   */
  const names = shDriver
    .getRange(
      1,
      1,
      lastRow,
      1
    )
    .getDisplayValues();

  names.forEach((row, index) => {
    const fullName = cleanText_(row[0]);

    if (!fullName) {
      return;
    }

    const parts =
      getCleanNameParts_(fullName);

    const firstName =
      normalizeKey_(parts[0] || "");

    if (!firstName) {
      return;
    }

    /*
     * Ignora posibles encabezados en la fila 1.
     */
    if (
      index === 0 &&
      (
        firstName === "driver" ||
        firstName === "name" ||
        firstName === "driver name"
      )
    ) {
      return;
    }

    counts.set(
      firstName,
      (counts.get(firstName) || 0) + 1
    );
  });

  return counts;
}

/**
 * Crea el nombre que se escribirá
 * en a&aprueba columna B.
 *
 * La repetición se determina usando
 * la columna A de DRIVER.
 *
 * EJEMPLOS:
 *
 * Un solo John:
 * John Kent -> JOHN
 *
 * Clarence repetidos:
 * Clarence Wendell Sneed
 * -> CLARENCE S
 *
 * Clarence Gregory Taylor
 * -> CLARENCE T
 */
function crearNombreDriverSalida_(
  driverFull,
  firstNameCounts
) {
  const parts =
    getCleanNameParts_(driverFull);

  const firstName = parts[0] || "";

  if (!firstName) {
    return "";
  }

  let result =
    firstName.toUpperCase();

  const firstNameKey =
    normalizeKey_(firstName);

  const repeated =
    (
      firstNameCounts.get(
        firstNameKey
      ) || 0
    ) > 1;

  /*
   * Si el primer nombre solamente aparece una vez
   * en DRIVER, devuelve solo el primer nombre.
   *
   * John Kent -> JOHN
   */
  if (!repeated) {
    return result;
  }

  let nameForInitial = "";

  /*
   * Si tiene 3 partes o más,
   * utiliza la tercera parte.
   *
   * Clarence Wendell Sneed
   *
   * parts[0] = Clarence
   * parts[1] = Wendell
   * parts[2] = Sneed
   *
   * Resultado: CLARENCE S
   */
  if (parts.length >= 3) {
    nameForInitial = parts[2];
  }

  /*
   * Si solo tiene 2 partes,
   * utiliza la segunda.
   *
   * John Kent -> JOHN K
   *
   * Esto solo ocurre si John está repetido
   * en la columna A de DRIVER.
   */
  else if (parts.length === 2) {
    nameForInitial = parts[1];
  }

  if (nameForInitial) {
    result +=
      ` ${nameForInitial
        .charAt(0)
        .toUpperCase()}`;
  }

  return result;
}

/**
 * Limpia el nombre del driver.
 *
 * Ignora cualquier identificador que empiece
 * desde una parte que contenga "(".
 *
 * Ejemplo:
 *
 * Clarence Wendell Sneed (AAXI123)
 *
 * Se convierte en:
 *
 * [
 *   "Clarence",
 *   "Wendell",
 *   "Sneed"
 * ]
 */
function getCleanNameParts_(name) {
  const rawParts = String(name || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);

  const parenIndex =
    rawParts.findIndex(
      part => part.includes("(")
    );

  const nameParts =
    parenIndex >= 0
      ? rawParts.slice(0, parenIndex)
      : rawParts;

  return nameParts
    .map(part =>
      part.replace(
        /[^A-Za-zÀ-ÖØ-öø-ÿ'’-]/g,
        ""
      )
    )
    .filter(Boolean);
}

/**
 * Extrae el número de la ruta.
 *
 * Ejemplos:
 *
 * CX115 -> 115
 * cx139 -> 139
 * 145   -> 145
 */
function extractRouteNumber_(routeRaw) {
  const s = cleanText_(routeRaw);

  const match = s.match(/(\d+)/);

  return match
    ? match[1]
    : s.replace(/CX/ig, "").trim();
}

/**
 * Elimina el prefijo STG
 * de la staging area.
 *
 * Ejemplo:
 *
 * STG.A.17 -> A.17
 */
function stripSTGPrefix_(areaRaw) {
  return cleanText_(areaRaw)
    .replace(
      /^STG[\s.\-]*/i,
      ""
    )
    .trim();
}

/**
 * Normaliza el texto de los sizes.
 */
function normalizeSizeText_(value) {
  return cleanText_(value)
    .toUpperCase()
    .replace(/[–—]/g, "-")
    .replace(/\s+/g, " ");
}

/**
 * Convierte cualquier valor
 * en texto limpio.
 */
function cleanText_(value) {
  return String(
    value == null ? "" : value
  ).trim();
}

/**
 * Normaliza valores para compararlos.
 */
function normalizeKey_(value) {
  return cleanText_(value)
    .toLowerCase()
    .replace(/\s+/g, " ");
}

/**
 * Compara dos horas.
 */
function compareTimes_(a, b) {
  return (
    parseTimeToMinutes_(a) -
    parseTimeToMinutes_(b)
  );
}

/**
 * Convierte una hora a minutos.
 *
 * Acepta:
 *
 * 11:20 AM
 * 12:00 PM
 * 23:20
 */
function parseTimeToMinutes_(value) {
  const t =
    cleanText_(value).toUpperCase();

  const amPmMatch = t.match(
    /^(\d{1,2}):(\d{2})\s*(AM|PM)$/
  );

  if (amPmMatch) {
    let hour =
      Number(amPmMatch[1]);

    const minute =
      Number(amPmMatch[2]);

    const period =
      amPmMatch[3];

    if (
      period === "AM" &&
      hour === 12
    ) {
      hour = 0;
    }

    if (
      period === "PM" &&
      hour !== 12
    ) {
      hour += 12;
    }

    return hour * 60 + minute;
  }

  const twentyFourHourMatch = t.match(
    /^(\d{1,2}):(\d{2})$/
  );
