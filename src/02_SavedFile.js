const labHeaders = {
  researchLevel: "researchLevel",
}

const workshopHeaders = {
  presetNames: "workshopPresetName",
  upgradeAttackLevels: "upgradeWorkshopLevel",
  upgradeDefenseLevels: "upgradeWorkshopDefenseLevel",
  upgradeUtilityLevels: "upgradeWorkshopUtilityLevel",
  presetUpgradeAttackLevels: "presetUpgradeWorkshopLevel",
  presetUpgradeDefenseLevels: "presetUpgradeWorkshopDefenseLevel",
  presetUpgradeUtilityLevels: "presetUpgradeWorkshopUtilityLevel",
  enhancementAttackLevels: "enhancementLevel",
  enhancementDefenseLevels: "enhancementDefenseLevel",
  enhancementUtilityLevels: "enhancementUtilityLevel",
  presetEnhancementAttackLevels: "presetEnhancementLevel",
  presetEnhancementDefenseLevels: "presetEnhancementDefenseLevel",
  presetEnhancementUtilityLevels: "presetEnhancementUtilityLevel",
  upgradeAttackUnlocked: "upgradeTierUnlocked",
  upgradeDefenseUnlocked: "upgradeDefenseTierUnlocked",
  upgradeUtilityUnlocked: "upgradeUtilityTierUnlocked",
  presetUpgradeAttackUnlocked: "presetUpgradeTierUnlocked",
  presetUpgradeDefenseUnlocked: "presetUpgradeDefenseTierUnlocked",
  presetUpgradeUtilityUnlocked: "presetUpgradeUtilityTierUnlocked",
}

const ultimateWeaponHeaders = {
  ultimateWeaponLevel: "ultimateWeaponLevel",
  ultimateWeaponUnlocked: "ultimateWeaponUnlocked",
  ultimateWeaponPlusLevel: "ultimateWeaponPlusLevel",
  ultimateWeaponPlusUnlocked: "ultimateWeaponPlusUnlocked",
}

const themesHeaders = {
  towerSkins: "towerUnlocked",
  backgroundSkins: "backgroundUnlocked",
  menuSkins: "menuUnlocked",
  guardianSkins: "guardianSkinUnlocked",
  profileBanners: "profileBannerUnlocked",
  songs: "trackAvailable",
}

const botHeaders = {
  presetNames: "botPresetName",
  flameBotPresets: "flameBotPresets",
  thunderBotPresets: "thunderBotPresets",
  goldenBotPresets: "goldenBotPresets",
  amplifyBotPresets: "amplifyBotPresets",
  botBotPresets: "botBotPresets",
  synchronicityPresets: "synchronicityPresets",
}

const relicHeaders = {
  relicsUnlocked: "relicsUnlocked",
}

const vaultHeaders = {
  harmonyNodesUnlocked: "harmonyNodesUnlocked",
  powerNodesUnlocked: "powerNodesUnlocked",
  powerNodesLevel: "powerNodesLevel",
}

const cardsHeaders = {
  cardLevel: "cardLevel",
  cardMasteryUnlocked: "cardMasteryUnlocked",
  presetNames: "presetName",
  presetSlots: "slotPresetCardAssignedBool",
  presetCards: "slotPresetCardInt",
  slotsUnlocked: "slotsUnlocked",
}

const moduleHeaders = {
  moduleEquipped: "moduleEquipped",
  inventory: "inventory",
  assistModuleSlots: "assistModuleSlots",
}

const guardianHeaders = {
  // guardianChipSlot: "guardianChipSlot",
  guardianChipUnlocked: "guardianChipUnlocked",
  guardianChipLevel: "guardianChipLevel",
}

const PlayerStuffHeaders = {
  playerID: "playfabID",
  currentTier: "currentTier",
  tourneyID: "leagueID",
  addPack: "disableAdsUnlockedBool",
  starterPack: "starterPackUnlockedBool",
  epicPack: "epicPackUnlockedBool",
  highestWavePerTier: "highestWaveThisTier",
  premiumPass: "milestonesPremiumUnlocked",
  atkDissonance: "dissonanceDamageBoost",
  hpDissonance: "dissonanceHealthBoost",
  coinDissonance: "dissonanceCoinBoost",
  uwDissonance: "dissonanceUltDamageBoost",
}

function parseSaveFileBytes(byteArray, fileName, showAll, sheetIDs) {
  const uint8 = Uint8Array.from(byteArray);
  // Decompress the GZIP data
  const decompressedBlob = Utilities.ungzip(
    Utilities.newBlob(Array.from(uint8), "application/x-gzip"),
  );
  // Convert the decompressed Blob to a Uint8Array
  const bytes = blobToUint8Array_(decompressedBlob);
  // Parse the NRBF data
  // .NET's BinaryFormatter uses a format called NRBF (Net Remote Binary Format) to serialize objects. This format is not natively supported in JavaScript, so we need to implement a parser for it.
  // The parseNRBF function is responsible for parsing the NRBF data and returning a JavaScript object representation of the serialized data.
  // It does both string length prefixing and UTF-8 decoding, which are necessary for correctly interpreting the serialized data.
  const data = parseNRBF(bytes);

  sheetIDs = sheetIDs || {};
  
  function extratctDataByHeaders(headers) {
    var values = {};
    for (const sheetKey of Object.keys(headers)) {
      const saveKey = headers[sheetKey];
      values[sheetKey] = data.hasOwnProperty(saveKey) ? data[saveKey] : null;
    }
    return values;
  }

  // Extract lab data
  var labValues = extratctDataByHeaders(labHeaders);
  var laboratoryData = lab.parseLabData(labValues);

  // Extract workshop data
  var workshopValues = extratctDataByHeaders(workshopHeaders);
  var workshopData = workshop.parseWorkshopData(workshopValues);

  // Extract ultimate weapon data
  var ultimateWeaponValues = extratctDataByHeaders(ultimateWeaponHeaders);
  var ultimateWeaponData = ultimate.parseUltimateWeaponData(ultimateWeaponValues);

  // Extract Themes & Songs data
  var themesValues = extratctDataByHeaders(themesHeaders);
  var themesData = themes.parseThemesData(
    themesValues,
    sheetIDs["Themes & Songs"],
  );

  // Extract bot data
  var botValues = extratctDataByHeaders(botHeaders);
  var botData = bots.parseBotsData(botValues);

  // Extract relic data
  var relicValues = extratctDataByHeaders(relicHeaders);
  var relicData = relics.parseRelicsData(relicValues, sheetIDs["Relics"]);

  // Extract vault data
  var vaultValues = extratctDataByHeaders(vaultHeaders);
  var vaultData = vault.parseVaultData(vaultValues);

  // Extract Cards data
  var cardsValues = extratctDataByHeaders(cardsHeaders);
  var cardsData = cards.parseCardsData(cardsValues);

  //Extract Modules data
  var moduleValues = extratctDataByHeaders(moduleHeaders);
  var moduleData = modules.parseModulesData(moduleValues);

  // Extract Guardian data
  var guardianValues = extratctDataByHeaders(guardianHeaders);
  var guardianData = guardians.parseGuardiansData(guardianValues);

  // Extract Player & Stuff data
  var playerStuffValues = extratctDataByHeaders(PlayerStuffHeaders);
  var playerStuffdata = playerStuff.parsePlayerStuffData(playerStuffValues);


  const parsed = {
    "Laboratory": laboratoryData,
    "Workshop": workshopData,
    "Ultimate Weapon": ultimateWeaponData,
    "Themes & Songs": themesData,
    "Bots": botData,
    "Relics": relicData,
    "Vault": vaultData,
    "Cards": cardsData,
    "Modules": moduleData,
    "Guardians": guardianData,
    "Player & Stuff": playerStuffdata,
  };

  return {
    parsed: parsed,
    order: Object.keys(parsed),
  };
}

function blobToUint8Array_(blob) {
  const signedBytes = blob.getBytes();
  const out = new Uint8Array(signedBytes.length);
  for (let i = 0; i < signedBytes.length; i++) {
    out[i] = signedBytes[i] & 0xff;
  }
  return out;
}

function bigIntJsonReplacer_(key, value) {
  return typeof value === "bigint" ? value.toString() : value;
}

const RecordType = {
  SerializationHeader: 0,
  ClassWithId: 1,
  SystemClassWithMembers: 2,
  ClassWithMembers: 3,
  SystemClassWithMembersAndTypes: 4,
  ClassWithMembersAndTypes: 5,
  BinaryObjectString: 6,
  BinaryArray: 7,
  MemberPrimitiveTyped: 8,
  MemberReference: 9,
  ObjectNull: 10,
  MessageEnd: 11,
  BinaryLibrary: 12,
  ObjectNullMultiple256: 13,
  ObjectNullMultiple: 14,
  ArraySinglePrimitive: 15,
  ArraySingleObject: 16,
  ArraySingleString: 17,
};

const BinaryType = {
  Primitive: 0,
  String: 1,
  Object: 2,
  SystemClass: 3,
  Class: 4,
  ObjectArray: 5,
  StringArray: 6,
  PrimitiveArray: 7,
};

const PrimitiveType = {
  Boolean: 1,
  Byte: 2,
  Char: 3,
  Decimal: 5,
  Double: 6,
  Int16: 7,
  Int32: 8,
  Int64: 9,
  SByte: 10,
  Single: 11,
  TimeSpan: 12,
  DateTime: 13,
  UInt16: 14,
  UInt32: 15,
  UInt64: 16,
  Null: 17,
  String: 18,
};

const BinaryArrayType = {
  Single: 0,
  Jagged: 1,
  Rectangular: 2,
  SingleOffset: 3,
  JaggedOffset: 4,
  RectangularOffset: 5,
};

class ParseError extends Error {}

class NRBFParser {
  constructor(bytes) {
    this.buf = bytes;
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    this.pos = 0;
    this.objects = new Map();
    this.libraries = new Map();
    this.classDefs = new Map();
  }

  u8() {
    const v = this.view.getUint8(this.pos);
    this.pos += 1;
    return v;
  }
  i8() {
    const v = this.view.getInt8(this.pos);
    this.pos += 1;
    return v;
  }
  u16() {
    const v = this.view.getUint16(this.pos, true);
    this.pos += 2;
    return v;
  }
  i16() {
    const v = this.view.getInt16(this.pos, true);
    this.pos += 2;
    return v;
  }
  i32() {
    const v = this.view.getInt32(this.pos, true);
    this.pos += 4;
    return v;
  }
  u32() {
    const v = this.view.getUint32(this.pos, true);
    this.pos += 4;
    return v;
  }
  i64() {
    const v = this.view.getBigInt64(this.pos, true);
    this.pos += 8;
    return v;
  }
  u64() {
    const v = this.view.getBigUint64(this.pos, true);
    this.pos += 8;
    return v;
  }
  f32() {
    const v = this.view.getFloat32(this.pos, true);
    this.pos += 4;
    return v;
  }
  f64() {
    const v = this.view.getFloat64(this.pos, true);
    this.pos += 8;
    return v;
  }

  lps() {
    let length = 0,
      shift = 0,
      b;
    do {
      b = this.u8();
      length |= (b & 0x7f) << shift;
      shift += 7;
    } while (b & 0x80);
    const bytes = this.buf.subarray(this.pos, this.pos + length);
    this.pos += length;
    return utf8Decode(bytes);
  }

  readChar() {
    const b = this.u8();
    if (b < 0x80) return String.fromCharCode(b);
    if (b < 0xe0) {
      const b2 = this.u8();
      return String.fromCharCode(((b & 0x1f) << 6) | (b2 & 0x3f));
    }
    if (b < 0xf0) {
      const b2 = this.u8(),
        b3 = this.u8();
      return String.fromCharCode(
        ((b & 0x0f) << 12) | ((b2 & 0x3f) << 6) | (b3 & 0x3f),
      );
    }
    const b2 = this.u8(),
      b3 = this.u8(),
      b4 = this.u8();
    const cp =
      ((b & 0x07) << 18) |
      ((b2 & 0x3f) << 12) |
      ((b3 & 0x3f) << 6) |
      (b4 & 0x3f);
    return String.fromCodePoint(cp);
  }

  primitive(ptype) {
    switch (ptype) {
      case PrimitiveType.Boolean:
        return this.u8() !== 0;
      case PrimitiveType.Byte:
        return this.u8();
      case PrimitiveType.Char:
        return this.readChar();
      case PrimitiveType.Decimal:
        return this.lps();
      case PrimitiveType.Double:
        return this.f64();
      case PrimitiveType.Int16:
        return this.i16();
      case PrimitiveType.Int32:
        return this.i32();
      case PrimitiveType.Int64:
        return this.i64();
      case PrimitiveType.SByte:
        return this.i8();
      case PrimitiveType.Single:
        return this.f32();
      case PrimitiveType.TimeSpan:
        return this.i64();
      case PrimitiveType.DateTime:
        return this.u64();
      case PrimitiveType.UInt16:
        return this.u16();
      case PrimitiveType.UInt32:
        return this.u32();
      case PrimitiveType.UInt64:
        return this.u64();
      default:
        throw new ParseError(`Unhandled primitive type ${ptype}`);
    }
  }

  readBinaryLibrary() {
    const libId = this.i32();
    this.libraries.set(libId, this.lps());
  }

  readClassInfo() {
    const objId = this.i32();
    const name = this.lps();
    const count = this.i32();
    const memberNames = [];
    for (let i = 0; i < count; i++) memberNames.push(this.lps());
    return { objId, name, memberNames };
  }

  readMemberTypeInfo(count) {
    const binaryTypes = [];
    for (let i = 0; i < count; i++) binaryTypes.push(this.u8());
    const additional = [];
    for (const bt of binaryTypes) {
      if (bt === BinaryType.Primitive || bt === BinaryType.PrimitiveArray) {
        additional.push(this.u8());
      } else if (bt === BinaryType.SystemClass) {
        additional.push(this.lps());
      } else if (bt === BinaryType.Class) {
        const className = this.lps();
        const libId = this.i32();
        additional.push([className, libId]);
      } else {
        additional.push(null);
      }
    }
    return { binaryTypes, additional };
  }

  readClassWithMembersAndTypes() {
    const { objId, name, memberNames } = this.readClassInfo();
    const { binaryTypes, additional } = this.readMemberTypeInfo(
      memberNames.length,
    );
    this.i32();
    this.classDefs.set(objId, { name, memberNames, binaryTypes, additional });
    const values = this.readMemberValues(binaryTypes, additional);
    const obj = {};
    memberNames.forEach((n, i) => (obj[n] = values[i]));
    obj._class = name;
    this.objects.set(objId, obj);
    return objId;
  }

  readSystemClassWithMembersAndTypes() {
    const { objId, name, memberNames } = this.readClassInfo();
    const { binaryTypes, additional } = this.readMemberTypeInfo(
      memberNames.length,
    );
    this.classDefs.set(objId, { name, memberNames, binaryTypes, additional });
    const values = this.readMemberValues(binaryTypes, additional);
    const obj = {};
    memberNames.forEach((n, i) => (obj[n] = values[i]));
    obj._class = name;
    this.objects.set(objId, obj);
    return objId;
  }

  readClassWithId() {
    const objId = this.i32();
    const metadataId = this.i32();
    const def = this.classDefs.get(metadataId);
    const values = this.readMemberValues(def.binaryTypes, def.additional);
    const obj = {};
    def.memberNames.forEach((n, i) => (obj[n] = values[i]));
    obj._class = def.name;
    this.objects.set(objId, obj);
    return objId;
  }

  readMemberValues(binaryTypes, additional) {
    const out = [];
    for (let i = 0; i < binaryTypes.length; i++) {
      out.push(this.readValue(binaryTypes[i], additional[i]));
    }
    return out;
  }

  readValue(bt, additional) {
    if (bt === BinaryType.Primitive) return this.primitive(additional);
    return this.readInlineValue();
  }

  readInlineValue() {
    if (this.pos >= this.buf.length) return null;
    const rec = this.u8();
    switch (rec) {
      case RecordType.BinaryObjectString:
        return this.readBinaryObjectStringBody();
      case RecordType.MemberReference:
        return { _ref: this.i32() };
      case RecordType.ObjectNull:
        return null;
      case RecordType.ObjectNullMultiple256:
      case RecordType.ObjectNullMultiple:
        throw new ParseError(
          `${rec} is only valid inside an array, not as a single member value`,
        );
      case RecordType.ClassWithMembersAndTypes:
        return { _ref: this.readClassWithMembersAndTypes() };
      case RecordType.SystemClassWithMembersAndTypes:
        return { _ref: this.readSystemClassWithMembersAndTypes() };
      case RecordType.ClassWithId:
        return { _ref: this.readClassWithId() };
      case RecordType.ArraySinglePrimitive:
        return this.readArraySinglePrimitive();
      case RecordType.ArraySingleObject:
        return this.readArraySingleObject();
      case RecordType.ArraySingleString:
        return this.readArraySingleString();
      case RecordType.BinaryArray:
        return this.readBinaryArray();
      case RecordType.MemberPrimitiveTyped:
        return this.primitive(this.u8());
      case RecordType.BinaryLibrary:
        this.readBinaryLibrary();
        return this.readInlineValue();
      default:
        throw new ParseError(
          `Unexpected record type ${rec} at offset ${this.pos}`,
        );
    }
  }

  readBinaryObjectStringBody() {
    const objId = this.i32();
    const s = this.lps();
    this.objects.set(objId, s);
    return s;
  }

  readArraySinglePrimitive() {
    const objId = this.i32();
    const length = this.i32();
    const ptype = this.u8();
    const arr = [];
    for (let i = 0; i < length; i++) arr.push(this.primitive(ptype));
    this.objects.set(objId, arr);
    return arr;
  }

  readArraySingleString() {
    const objId = this.i32();
    const length = this.i32();
    const arr = this.readArrayElements(length);
    this.objects.set(objId, arr);
    return arr;
  }

  readArraySingleObject() {
    const objId = this.i32();
    const length = this.i32();
    const arr = this.readArrayElements(length);
    this.objects.set(objId, arr);
    return arr;
  }

  readArrayElements(length) {
    const arr = [];
    let i = 0;
    while (i < length) {
      const pos = this.pos;
      if (this.pos >= this.buf.length) break;
      const rec = this.u8();
      if (rec === RecordType.ObjectNull) {
        arr.push(null);
        i += 1;
      } else if (rec === RecordType.ObjectNullMultiple256) {
        const count = this.u8();
        for (let k = 0; k < count; k++) arr.push(null);
        i += count;
      } else if (rec === RecordType.ObjectNullMultiple) {
        const count = this.i32();
        for (let k = 0; k < count; k++) arr.push(null);
        i += count;
      } else {
        this.pos = pos;
        arr.push(this.readInlineValue());
        i += 1;
      }
    }
    return arr;
  }

  readBinaryArray() {
    const objId = this.i32();
    const arrayType = this.u8();
    const rank = this.i32();
    const lengths = [];
    for (let i = 0; i < rank; i++) lengths.push(this.i32());
    if (
      arrayType === BinaryArrayType.SingleOffset ||
      arrayType === BinaryArrayType.JaggedOffset ||
      arrayType === BinaryArrayType.RectangularOffset
    ) {
      for (let i = 0; i < rank; i++) this.i32();
    }
    const bt = this.u8();
    let additional = null;
    if (bt === BinaryType.Primitive || bt === BinaryType.PrimitiveArray) {
      additional = this.u8();
    } else if (bt === BinaryType.SystemClass) {
      additional = this.lps();
    } else if (bt === BinaryType.Class) {
      additional = [this.lps(), this.i32()];
    }

    let total = 1;
    for (const d of lengths) total *= d;

    let arr;
    if (bt === BinaryType.Primitive && additional !== null) {
      arr = [];
      for (let i = 0; i < total; i++) arr.push(this.primitive(additional));
    } else {
      arr = this.readArrayElements(total);
    }
    this.objects.set(objId, arr);
    return arr;
  }
  parse() {
    const firstByte = this.u8();
    if (firstByte !== RecordType.SerializationHeader) {
      throw new ParseError(
        `Stream does not start with a SerializationHeaderRecord (got byte 0x${firstByte.toString(16)})`,
      );
    }
    const rootId = this.i32();
    this.i32();
    this.i32();
    this.i32();

    while (true) {
      if (this.pos >= this.buf.length) break;
      const pos = this.pos;
      const rec = this.u8();
      if (rec === RecordType.MessageEnd) break;
      else if (rec === RecordType.BinaryLibrary) this.readBinaryLibrary();
      else if (rec === RecordType.ClassWithMembersAndTypes)
        this.readClassWithMembersAndTypes();
      else if (rec === RecordType.SystemClassWithMembersAndTypes)
        this.readSystemClassWithMembersAndTypes();
      else if (rec === RecordType.ClassWithId) this.readClassWithId();
      else if (rec === RecordType.BinaryObjectString)
        this.readBinaryObjectStringBody();
      else if (rec === RecordType.ArraySinglePrimitive)
        this.readArraySinglePrimitive();
      else if (rec === RecordType.ArraySingleObject)
        this.readArraySingleObject();
      else if (rec === RecordType.ArraySingleString)
        this.readArraySingleString();
      else if (rec === RecordType.BinaryArray) this.readBinaryArray();
      else
        throw new ParseError(
          `Unexpected top-level record ${rec} at offset ${pos}`,
        );
    }

    return this.resolve(this.objects.get(rootId));
  }

  resolve(value, seen) {
    seen = seen || new Set();
    if (
      value &&
      typeof value === "object" &&
      !Array.isArray(value) &&
      !(value instanceof Uint8Array)
    ) {
      const keys = Object.keys(value);
      if (keys.length === 1 && keys[0] === "_ref") {
        const refId = value._ref;
        if (!this.objects.has(refId))
          throw new ParseError(`Unresolved object reference: ID ${refId}`);
        return this.resolve(this.objects.get(refId), seen);
      }
      const out = {};
      for (const k of keys) out[k] = this.resolve(value[k], seen);
      return unwrapCollection(out);
    }
    if (Array.isArray(value)) {
      return value.map((v) => this.resolve(v, seen));
    }
    return value;
  }
}

function unwrapCollection(obj) {
  const keys = Object.keys(obj);
  if (
    keys.length === 2 &&
    keys.includes("value__") &&
    keys.includes("_class") &&
    typeof obj.value__ === "number"
  ) {
    return obj.value__;
  }
  const cls = obj._class || "";
  if (cls.startsWith("System.Collections.Generic.List`1")) {
    const items = obj._items || [];
    const size = typeof obj._size === "number" ? obj._size : items.length;
    return items.slice(0, size);
  }
  if (cls.startsWith("System.Collections.Generic.Dictionary`2")) {
    const pairs = obj.KeyValuePairs || [];
    const out = {};
    for (const p of pairs) {
      if (p && typeof p === "object" && "key" in p) out[p.key] = p.value;
    }
    return out;
  }
  if (cls.startsWith("System.Collections.Generic.KeyValuePair`2")) {
    return { key: obj.key, value: obj.value };
  }
  if (cls.startsWith("System.Collections.Generic.")) {
    const useful = {};
    let any = false;
    for (const k of Object.keys(obj)) {
      if (k !== "_class" && !k.startsWith("_")) {
        useful[k] = obj[k];
        any = true;
      }
    }
    if (!any) return null;
  }
  return obj;
}

function utf8Decode(bytes) {
  if (typeof TextDecoder !== "undefined") {
    return new TextDecoder("utf-8").decode(bytes);
  }
  let out = "",
    i = 0;
  while (i < bytes.length) {
    const b1 = bytes[i++];
    if (b1 < 0x80) {
      out += String.fromCharCode(b1);
      continue;
    }
    if (b1 < 0xe0) {
      const b2 = bytes[i++];
      out += String.fromCharCode(((b1 & 0x1f) << 6) | (b2 & 0x3f));
      continue;
    }
    if (b1 < 0xf0) {
      const b2 = bytes[i++],
        b3 = bytes[i++];
      out += String.fromCharCode(
        ((b1 & 0x0f) << 12) | ((b2 & 0x3f) << 6) | (b3 & 0x3f),
      );
      continue;
    }
    const b2 = bytes[i++],
      b3 = bytes[i++],
      b4 = bytes[i++];
    out += String.fromCodePoint(
      ((b1 & 0x07) << 18) |
        ((b2 & 0x3f) << 12) |
        ((b3 & 0x3f) << 6) |
        (b4 & 0x3f),
    );
  }
  return out;
}

function parseNRBF(bytes) {
  return new NRBFParser(bytes).parse();
}
