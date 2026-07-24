import React, { useState, useMemo, useRef, useCallback, useEffect } from "react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LineChart, Line, PieChart, Pie, Cell } from "recharts";
import * as XLSX from "xlsx";
import { cloudEnabled, subscribeReport, saveReport } from "./cloud";

// ── Persistence: remember the last uploaded data across page reloads ──────────
// Both upload channels (sales data + area/machine info) are saved to
// localStorage and restored on load. New uploads overwrite the saved copy.
const LS = {
  get(key, fallback) {
    try { const v = localStorage.getItem(key); return v == null ? fallback : JSON.parse(v); }
    catch (e) { return fallback; }
  },
  set(key, val) {
    try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) {}
  },
};

// ── Data ──────────────────────────────────────────────────────────────────────
// Default reference data (replaced when user uploads machine info file)
const DEFAULT_INSTALL = {"IG01690002":"17-01-26","IG01690003":"17-01-26","IG01690004":"17-01-26","IG01690006":"17-01-26","IG01690005":"18-01-26","IG01690011":"18-01-26","IG01690001":"20-01-26","IG01690012":"20-01-26","IG01690013":"21-01-26","IG01690014":"21-01-26","IG01690015":"21-01-26","IG01690009":"25-01-26","IG01690007":"22-01-26","IG01690008":"27-01-26","IG01690010":"23-01-26","IG01690017":"23-01-26","IG01690018":"24-01-26","IG01690031":"12-02-26","IG01690028":"15-02-26","IG01690029":"12-02-26","IG01690032":"11-02-26","IG01690020":"11-02-26","IG01690019":"11-02-26","IG01690030":"11-02-26","IG01690034":"12-02-26","IG01690025":"12-02-26","IG01690033":"12-02-26","IG01690022":"21-02-26","IG01690035":"17-02-26","IG01690039":"18-02-26","IG01690021":"18-02-26","IG01690037":"21-02-26","IG01690038":"17-02-26","IG01690040":"25-02-26","IG01690036":"17-02-26","IG01690016":"19-02-26","IG01690023":"19-02-26","IG01690024":"19-02-26","IG01690027":"21-02-26","IG01690026":"21-02-26"};
const DEFAULT_RENT = {"IG01690002":1050,"IG01690003":1200,"IG01690004":1200,"IG01690006":1050,"IG01690005":1150,"IG01690011":1050,"IG01690001":1050,"IG01690012":1050,"IG01690013":1200,"IG01690014":1050,"IG01690015":1050,"IG01690009":1150,"IG01690007":1050,"IG01690008":1050,"IG01690010":1200,"IG01690017":1200,"IG01690018":1050,"IG01690031":1050,"IG01690028":1050,"IG01690029":1050,"IG01690032":1050,"IG01690020":1050,"IG01690019":1050,"IG01690030":1250,"IG01690034":1050,"IG01690025":1334,"IG01690033":1334,"IG01690022":1050,"IG01690035":1334,"IG01690039":1334,"IG01690021":1334,"IG01690037":1334,"IG01690038":1050,"IG01690040":1334,"IG01690036":1050,"IG01690016":1550,"IG01690023":1550,"IG01690024":1050,"IG01690027":1050,"IG01690026":1550};
const DEFAULT_LOCATION = {"IG01690002":"หอพัก ถนอมสุข","IG01690003":"โครงการลัคกี้เพลสทรัพย์พัฒนา","IG01690004":"โครงการไทยฟู๊ดเฟสมาร์เก็ตซับพัฒนา","IG01690006":"ค่ายทหารพหลโยธินซอย 2","IG01690005":"บจก.บิ๊กดีล โกลบอล (พระราม 3)","IG01690011":"ร้านเมืองใหม่ เคหะภัณฑ์","IG01690001":"Yori wash&dry รังสิต-คลอง2","IG01690012":"26 ซอย รณสิทธิพิชัย 8/3","IG01690013":"ตุสิตาธงฟ้าประชารัฐ 29/7 ม.11 ถ.ซอยพระราชวิริยาภรณ์ 16","IG01690014":"mix values ตำบล บางพูน อำเภอเมืองปทุมธานี","IG01690015":"เอื้ออาทร (บ้านเดี่ยว) ลาดหลุมแก้ว","IG01690009":"บจก.เอ็มเอสเค คอร์ปอเรชั่น (อ่อนนุชขาออก)","IG01690007":"สุวินทวงศ์ 13","IG01690008":"ร้านสะดวกซัก Wash Up Express (วอชอัพ เอ็กซ์เพรส)","IG01690010":"เมืองใหม่บางพลี 6","IG01690017":"ร้านนวดคุณอ้อย","IG01690018":"ร้านเสริมสวยกะตังค์","IG01690031":"ร้านทำผม อิสระภาพซอย 15","IG01690028":"โครงการ Fei market plaza ซ. วัดพระเงิน","IG01690029":"SriyenService มบ.เอเซียโฮม","IG01690032":"หมู่บ้านพงษ์ศิริชัย4","IG01690020":"จุรินทร์จัดสรร","IG01690019":"ร้านตามฟาร์มฝัน บางบัวทอง","IG01690030":"สุขเจริญ แมนชั่น(ทับช้าง)","IG01690034":"ร้าน Hatoxthegreenhob","IG01690025":"บิ๊กซีมินิ เทิดราชัน ซอย 1","IG01690033":"บิ๊กซีมินิ เทียนทะเล 28","IG01690022":"ร้านขายของชำ น้องวิน","IG01690035":"บิ๊กซีมินิ หมู่บ้านซื่อตรง","IG01690039":"บิ๊กซีมินิ แฉล้มนิมิตร","IG01690021":"บิ๊กซีมินิ พัฒนาชนบท 4, กรุงเทพฯ ลาดกระบัง","IG01690037":"บิ๊กซีมินิ สุขาภิบาล 5 ซอย 32, กรุงเทพ","IG01690038":"หน้าบ้านคุณวราสิทธิ์ ร้านมารวยอาหารสัตว์","IG01690040":"บิ๊กซีมินิ สุคนธสวัสดิ์","IG01690036":"Bear's house apartment","IG01690016":"ไทยฟู้ดส์ เฟรซ มาร์เก็ต สาขาชินเขต (งามวงศ์วาน 47)","IG01690023":"ไทยฟู้ดส์ เฟรซ มาร์เก็ต สาขาทีเค-ไทรน้อย นนทบุรี","IG01690024":"NP CARWASH ล้างรถหยอดเหรียญ","IG01690027":"หน้าบ้านคุณโค้ก หนองจอก","IG01690026":"ไทยฟู้ดส์ เฟรซ มาร์เก็ต สาขาตลาดทรัพย์มงคลคลอง 9"};
const DEFAULT_PHOTO = {"IG01690002":"https://drive.google.com/drive/folders/1wDZWyVZCbdhBGNYJGn02UX0XC00LTy_J?usp=drive_link","IG01690003":"https://drive.google.com/drive/folders/1kNqEuoWi2KUg_6BJYByTX_Ci3Q8nXcQU?usp=drive_link","IG01690004":"https://drive.google.com/drive/folders/1okws2MAYgbl81BAGSp51iKEouKZ94cqu?usp=drive_link","IG01690006":"https://drive.google.com/drive/folders/1OV9Opljw3n-wHyvkxVAVUUOZKufXXXIK?usp=drive_link","IG01690005":"https://drive.google.com/drive/folders/1vVWs8GOi5HiWC79_5ENi_bOHsRjKunsh?usp=drive_link","IG01690011":"https://drive.google.com/drive/folders/112Op3Dkz0ivrSVU-B3rvyJuQcZwdXKuE?usp=drive_link","IG01690001":"https://drive.google.com/drive/folders/14JWOYqfZo4Pq0jfs695w7rQlXuO31sMW?usp=drive_link","IG01690012":"https://drive.google.com/drive/folders/1_or9c2q43VsdKwvkHFmWZERLaPJek3Bz?usp=drive_link","IG01690013":"https://drive.google.com/drive/folders/1yC_8heIldkUcvTw071jeBZbx7nvPGc2L?usp=drive_link","IG01690014":"https://drive.google.com/drive/folders/1fxm2JvgnwkrP2j3XhbwKNs8aapo2mbTa?usp=drive_link","IG01690015":"https://drive.google.com/drive/folders/1xJw53iZTzgD169LeHVgwtPEhEs4fpevC?usp=drive_link","IG01690009":"https://drive.google.com/drive/folders/1Wskr93v11XJX_x3fJshEbpHjpSGu6lfK?usp=drive_link","IG01690007":"https://drive.google.com/drive/folders/15NW9HyFbvyTsiN4qpPd1AdNOkpkPzbQI?usp=drive_link","IG01690008":"https://drive.google.com/drive/folders/1TOZkkjLuvISASQywBPDERAHEunBJY9eN?usp=drive_link","IG01690010":"https://drive.google.com/drive/folders/1fL_F0MZN4b8p1x4TOC-x2NMMMnzew5Dk?usp=drive_link","IG01690017":"https://drive.google.com/drive/folders/1dXH6-DI38rGk9LzoXxZSk4R3xqpDXxyw?usp=drive_link","IG01690018":"https://drive.google.com/drive/folders/1O8VbC54Xr__3YNoP5cZtwzBNWL9W3PSO?usp=drive_link","IG01690031":"https://drive.google.com/drive/folders/17P1ocAliESi0QasRj7UKzq2J6OJXMpt2?usp=drive_link","IG01690028":"https://drive.google.com/drive/folders/1TcD32Bm5_zzF_xEAjXPeSOPPspwSm70_?usp=drive_link","IG01690029":"https://drive.google.com/drive/folders/1kfxHD-ijGAxr4XelI1ePsuHZ3_WDGAdI?usp=drive_link","IG01690032":"https://drive.google.com/drive/folders/1LHUCesSCAz3GkxDMBeBvJI80VaC3Mtf8?usp=drive_link","IG01690020":"https://drive.google.com/drive/folders/1VDlVwSqIe-yzJKzYcZwfXT_0HH6Uh7S-?usp=drive_link","IG01690019":"https://drive.google.com/drive/folders/1kopXLXZXG_c9j3syMJO6m7lgKq57ZakC?usp=drive_link","IG01690030":"https://drive.google.com/drive/folders/1C1Rea5P3l_FW1X2gJlh8S1Z0QWSDiG7P?usp=drive_link","IG01690034":"https://drive.google.com/drive/folders/1FYSP0kQrZxdvGdgnXUOgvxVcujr8bzp-?usp=drive_link","IG01690025":"https://drive.google.com/drive/folders/1YLnM7IYNYdBC-c7J7wQmqneaUjsvJjJ1?usp=drive_link","IG01690033":"https://drive.google.com/drive/folders/1oufglrauwFu_DUBAWmsgaNQd2pKyrrgz?usp=drive_link","IG01690022":"https://drive.google.com/drive/folders/1eB79mUU5Cc9s_m0sJQyhAMYhfP80lutc?usp=drive_link","IG01690035":"https://drive.google.com/drive/folders/15v6mLgZ2cKKnChWVRTeiqrx8FoN4RVv2?usp=drive_link","IG01690039":"https://drive.google.com/drive/folders/1wJL5VhjJTof1WusxFPqDKEzs3IZe9Fdg?usp=drive_link","IG01690021":"https://drive.google.com/drive/folders/1LLG2clSbvNT9Ltc4-yL3awuF9CNxDBYv?usp=drive_link","IG01690037":"https://drive.google.com/drive/folders/1PJOdYstdQ36uX-yB5JaLvmxdwxa6JrYQ?usp=drive_link","IG01690038":"https://drive.google.com/drive/folders/11IRme92EbuxLn764pz64M40CCMnw49IU?usp=drive_link","IG01690040":"https://drive.google.com/drive/folders/1E8lGb7HO4zbllcxnjAnUM7JF-IYNZVIM?usp=drive_link","IG01690036":"https://drive.google.com/drive/folders/1YR9e4-bdi1P5Z713d82305snBAgH1FQC?usp=drive_link","IG01690016":"https://drive.google.com/drive/folders/1hZjwJXR4S3jGjD1eFqjFfdhxr6EPlz2T?usp=drive_link","IG01690023":"https://drive.google.com/drive/folders/1QczAQ9WLRbmf5EG-bFjnY1cqez-Vhp0_?usp=drive_link","IG01690024":"https://drive.google.com/drive/folders/155g1NhPaGK-kVd5ASICtzP1YDEJZPC7C","IG01690027":"https://drive.google.com/drive/folders/1sjmyn30IpsX1GOrDYoQeEtjtVjj_t_h8?usp=drive_link","IG01690026":"https://drive.google.com/drive/folders/1WNb1cmB9xSJGW7inHLFqNrDatz88mavi?usp=drive_link"};
const DEFAULT_FOLDER_EMBED = {"IG01690002":"https://drive.google.com/embeddedfolderview?id=1wDZWyVZCbdhBGNYJGn02UX0XC00LTy_J#grid","IG01690003":"https://drive.google.com/embeddedfolderview?id=1kNqEuoWi2KUg_6BJYByTX_Ci3Q8nXcQU#grid","IG01690004":"https://drive.google.com/embeddedfolderview?id=1okws2MAYgbl81BAGSp51iKEouKZ94cqu#grid","IG01690006":"https://drive.google.com/embeddedfolderview?id=1OV9Opljw3n-wHyvkxVAVUUOZKufXXXIK#grid","IG01690005":"https://drive.google.com/embeddedfolderview?id=1vVWs8GOi5HiWC79_5ENi_bOHsRjKunsh#grid","IG01690011":"https://drive.google.com/embeddedfolderview?id=112Op3Dkz0ivrSVU-B3rvyJuQcZwdXKuE#grid","IG01690001":"https://drive.google.com/embeddedfolderview?id=14JWOYqfZo4Pq0jfs695w7rQlXuO31sMW#grid","IG01690012":"https://drive.google.com/embeddedfolderview?id=1_or9c2q43VsdKwvkHFmWZERLaPJek3Bz#grid","IG01690013":"https://drive.google.com/embeddedfolderview?id=1yC_8heIldkUcvTw071jeBZbx7nvPGc2L#grid","IG01690014":"https://drive.google.com/embeddedfolderview?id=1fxm2JvgnwkrP2j3XhbwKNs8aapo2mbTa#grid","IG01690015":"https://drive.google.com/embeddedfolderview?id=1xJw53iZTzgD169LeHVgwtPEhEs4fpevC#grid","IG01690009":"https://drive.google.com/embeddedfolderview?id=1Wskr93v11XJX_x3fJshEbpHjpSGu6lfK#grid","IG01690007":"https://drive.google.com/embeddedfolderview?id=15NW9HyFbvyTsiN4qpPd1AdNOkpkPzbQI#grid","IG01690008":"https://drive.google.com/embeddedfolderview?id=1TOZkkjLuvISASQywBPDERAHEunBJY9eN#grid","IG01690010":"https://drive.google.com/embeddedfolderview?id=1fL_F0MZN4b8p1x4TOC-x2NMMMnzew5Dk#grid","IG01690017":"https://drive.google.com/embeddedfolderview?id=1dXH6-DI38rGk9LzoXxZSk4R3xqpDXxyw#grid","IG01690018":"https://drive.google.com/embeddedfolderview?id=1O8VbC54Xr__3YNoP5cZtwzBNWL9W3PSO#grid","IG01690031":"https://drive.google.com/embeddedfolderview?id=17P1ocAliESi0QasRj7UKzq2J6OJXMpt2#grid","IG01690028":"https://drive.google.com/embeddedfolderview?id=1TcD32Bm5_zzF_xEAjXPeSOPPspwSm70_#grid","IG01690029":"https://drive.google.com/embeddedfolderview?id=1kfxHD-ijGAxr4XelI1ePsuHZ3_WDGAdI#grid","IG01690032":"https://drive.google.com/embeddedfolderview?id=1LHUCesSCAz3GkxDMBeBvJI80VaC3Mtf8#grid","IG01690020":"https://drive.google.com/embeddedfolderview?id=1VDlVwSqIe-yzJKzYcZwfXT_0HH6Uh7S-#grid","IG01690019":"https://drive.google.com/embeddedfolderview?id=1kopXLXZXG_c9j3syMJO6m7lgKq57ZakC#grid","IG01690030":"https://drive.google.com/embeddedfolderview?id=1C1Rea5P3l_FW1X2gJlh8S1Z0QWSDiG7P#grid","IG01690034":"https://drive.google.com/embeddedfolderview?id=1FYSP0kQrZxdvGdgnXUOgvxVcujr8bzp-#grid","IG01690025":"https://drive.google.com/embeddedfolderview?id=1YLnM7IYNYdBC-c7J7wQmqneaUjsvJjJ1#grid","IG01690033":"https://drive.google.com/embeddedfolderview?id=1oufglrauwFu_DUBAWmsgaNQd2pKyrrgz#grid","IG01690022":"https://drive.google.com/embeddedfolderview?id=1eB79mUU5Cc9s_m0sJQyhAMYhfP80lutc#grid","IG01690035":"https://drive.google.com/embeddedfolderview?id=15v6mLgZ2cKKnChWVRTeiqrx8FoN4RVv2#grid","IG01690039":"https://drive.google.com/embeddedfolderview?id=1wJL5VhjJTof1WusxFPqDKEzs3IZe9Fdg#grid","IG01690021":"https://drive.google.com/embeddedfolderview?id=1LLG2clSbvNT9Ltc4-yL3awuF9CNxDBYv#grid","IG01690037":"https://drive.google.com/embeddedfolderview?id=1PJOdYstdQ36uX-yB5JaLvmxdwxa6JrYQ#grid","IG01690038":"https://drive.google.com/embeddedfolderview?id=11IRme92EbuxLn764pz64M40CCMnw49IU#grid","IG01690040":"https://drive.google.com/embeddedfolderview?id=1E8lGb7HO4zbllcxnjAnUM7JF-IYNZVIM#grid","IG01690036":"https://drive.google.com/embeddedfolderview?id=1YR9e4-bdi1P5Z713d82305snBAgH1FQC#grid","IG01690016":"https://drive.google.com/embeddedfolderview?id=1hZjwJXR4S3jGjD1eFqjFfdhxr6EPlz2T#grid","IG01690023":"https://drive.google.com/embeddedfolderview?id=1QczAQ9WLRbmf5EG-bFjnY1cqez-Vhp0_#grid","IG01690024":"https://drive.google.com/embeddedfolderview?id=155g1NhPaGK-kVd5ASICtzP1YDEJZPC7C#grid","IG01690027":"https://drive.google.com/embeddedfolderview?id=1sjmyn30IpsX1GOrDYoQeEtjtVjj_t_h8#grid","IG01690026":"https://drive.google.com/embeddedfolderview?id=1WNb1cmB9xSJGW7inHLFqNrDatz88mavi#grid"};
const DEFAULT_MAP = {"IG01690002":"https://maps.app.goo.gl/PzVzQtPG8sU9Qyu49","IG01690003":"https://maps.app.goo.gl/AE3CwZmkPCgZprMD9","IG01690004":"https://maps.app.goo.gl/2hbnsgQBodWKo3KE8","IG01690006":"https://maps.google.com/maps?q=13.774927,100.547989","IG01690005":"https://maps.app.goo.gl/GsRXoG3NqM5zeah99","IG01690011":"https://maps.app.goo.gl/fgrHkaiUFeGvbkmMA","IG01690001":"https://maps.app.goo.gl/Gr7M9kYeTRvSpzdu5","IG01690012":"https://maps.app.goo.gl/21jW4J2EK99JmZAr6","IG01690013":"https://maps.app.goo.gl/XJfANQnhzvcSBdqk7","IG01690014":"https://maps.app.goo.gl/yg43a8bPVnzyuvux9","IG01690015":"https://maps.app.goo.gl/3zdX7hAFBTwjbZrC7","IG01690009":"https://maps.app.goo.gl/5X17gfqpZYb1VGa26","IG01690007":"https://maps.app.goo.gl/ZjmK7QVYTZ2nayWJA","IG01690008":"https://maps.app.goo.gl/CgWx5mpRmXmns4Hu8","IG01690010":"https://maps.app.goo.gl/q7mdQZqak88vfXan9","IG01690017":"https://maps.app.goo.gl/9SzTe6ywEKvWrKFL7","IG01690018":"https://maps.app.goo.gl/6GdKnX8UTVHkFTHD6","IG01690031":"https://maps.app.goo.gl/jhtNMK69aFpTgqxs9","IG01690028":"https://maps.app.goo.gl/YBTaqBn82Mkzy1Pu5","IG01690029":"https://maps.app.goo.gl/M1yxtwrrYcndNfpr6","IG01690032":"https://maps.app.goo.gl/1mHaRqDzrAjTNPXj8","IG01690020":"https://maps.app.goo.gl/ydcmr4hTfULTF1n9A","IG01690019":"https://maps.app.goo.gl/31rwG82FroLcFWWB8","IG01690030":"https://maps.app.goo.gl/ZUhJjAz9wYRdedLo6","IG01690034":"https://maps.app.goo.gl/pjq5s6Dggpo36i2r7","IG01690025":"https://maps.app.goo.gl/ipy1mTxxd9jsdR5N7","IG01690033":"https://maps.app.goo.gl/AH3d8BhP4h3krosM9","IG01690022":"https://maps.app.goo.gl/KQgWYKzjyBKgzn446","IG01690035":"https://maps.app.goo.gl/tCkPME7rwKPsEVT7A","IG01690039":"https://maps.app.goo.gl/PQX2emZFTGPT5cCX8","IG01690021":"https://maps.app.goo.gl/nAR9kajuuWyPweiF7","IG01690037":"https://maps.app.goo.gl/Hu6LY8DBkUbPbXoX7","IG01690038":"https://maps.app.goo.gl/ag8aUZeyspTiz3Ha8","IG01690040":"https://maps.app.goo.gl/eBXsRU5VkEUjSY4A8","IG01690036":"https://maps.app.goo.gl/WAE11tCPEnRg4j42A","IG01690016":"https://maps.app.goo.gl/XbEtq5P2dH6su9wz6","IG01690023":"https://maps.app.goo.gl/tWMXwEJPcFT9kNkq8","IG01690024":"https://maps.app.goo.gl/jaJgUbXo7Pd6nfzbA","IG01690027":"https://maps.app.goo.gl/LsvT1R73y7e9RJgDA","IG01690026":"https://maps.app.goo.gl/qDhTKDzH8MJ1UEyt5"};

// ── Parse machine-info file (install date + rent) ─────────────────────────────
// Detects device-id column, install-date column, rent column automatically
function parseMachineInfo(rows) {
  // rows: array of arrays (already parsed from xlsx or csv)
  // Find header row
  let hi = -1;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const r = rows[i].map(c => String(c).toLowerCase().trim());
    // needs at least a device-like column + one more data column
    if (r.some(c => /device|serial|mac|เครื่อง|machine|id/.test(c))) { hi = i; break; }
  }
  if (hi < 0) throw new Error("ไม่พบ header row ในไฟล์ข้อมูลเครื่อง");

  const hdrs = rows[hi].map(c => String(c).toLowerCase().trim());

  // Find device ID column
  const cDev = hdrs.findIndex(h => /device.detail|device id|serial|mac|เครื่อง|machine id/.test(h))
             ?? hdrs.findIndex(h => /device|id/.test(h));

  // Find install date column
  const cInst = hdrs.findIndex(h => /install|ติดตั้ง|วันที่|date/.test(h));

  // Find rent column
  const cRent = hdrs.findIndex(h => /rent|ค่าเช่า|เช่า|fee|price/.test(h));

  if (cDev < 0) throw new Error("ไม่พบคอลัมน์หมายเลขเครื่อง (Device/Serial/MAC)");
  if (cInst < 0 && cRent < 0) throw new Error("ไม่พบคอลัมน์วันติดตั้ง หรือ ค่าเช่า");

  // Find map/url column once before the loop
  const cMap   = hdrs.findIndex(h => /^map$|^url$|^link$|^พิกัด$/.test(h));
  const cLoc   = hdrs.findIndex(h => /location|สถานที่|place|ชื่อ/.test(h));
  const cPhoto = hdrs.findIndex(h => /photo|รูป|image|picture|img|drive/.test(h));

  const install = {}, rent = {}, mapData = {}, locData = {}, photoData = {};
  for (let i = hi + 1; i < rows.length; i++) {
    const row = rows[i];
    const dev = String(row[cDev] || "").trim();
    if (!dev) continue;

    if (cInst >= 0 && row[cInst]) {
      // Normalize date: could be "2026-01-20", "20/01/26", "20-01-26", Excel serial number
      let raw = row[cInst];
      let dateStr = "";
      if (typeof raw === "number") {
        // Excel serial date → JS Date
        const d = new Date(Math.round((raw - 25569) * 86400 * 1000));
        const dd = String(d.getUTCDate()).padStart(2,"0");
        const mm = String(d.getUTCMonth()+1).padStart(2,"0");
        const yy = String(d.getUTCFullYear()).slice(-2);
        dateStr = `${dd}-${mm}-${yy}`;
      } else {
        raw = String(raw).trim();
        // Try various date formats
        const m1 = raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); // YYYY-MM-DD
        const m2 = raw.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/); // DD-MM-YY or DD-MM-YYYY
        if (m1) {
          dateStr = `${m1[3].padStart(2,"0")}-${m1[2].padStart(2,"0")}-${String(m1[1]).slice(-2)}`;
        } else if (m2) {
          dateStr = `${m2[1].padStart(2,"0")}-${m2[2].padStart(2,"0")}-${String(m2[3]).slice(-2)}`;
        } else {
          dateStr = raw;
        }
      }
      install[dev] = dateStr;
    }

    if (cRent >= 0 && row[cRent]) {
      const r = parseFloat(String(row[cRent]).replace(/[^0-9.]/g,""));
      if (!isNaN(r)) rent[dev] = r;
    }

    if (cMap >= 0 && row[cMap]) {
      const u = String(row[cMap]).trim();
      if (u.startsWith("http")) mapData[dev] = u;
    }
    if (cLoc >= 0 && row[cLoc]) locData[dev] = String(row[cLoc]).trim();
    if (cPhoto >= 0 && row[cPhoto]) {
      const u = String(row[cPhoto]).trim();
      if (u.startsWith("http")) photoData[dev] = u;
    }
  }
  return { install, rent, mapData, locData, photoData };
}
const INIT_DATA = {"IG01690001":{"2026-01":{"total":113,"count":35,"cash":102,"qr":11},"2026-02":{"total":338,"count":60,"cash":328,"qr":10},"2026-03":{"total":118,"count":15,"cash":118,"qr":0}},"IG01690002":{"2026-01":{"total":435,"count":98,"cash":426,"qr":9},"2026-02":{"total":882,"count":173,"cash":851,"qr":31},"2026-03":{"total":134,"count":23,"cash":134,"qr":0}},"IG01690003":{"2026-01":{"total":137,"count":51,"cash":131,"qr":6},"2026-02":{"total":228,"count":94,"cash":205,"qr":23},"2026-03":{"total":63,"count":22,"cash":53,"qr":10}},"IG01690004":{"2026-01":{"total":304,"count":61,"cash":261,"qr":43},"2026-02":{"total":426,"count":100,"cash":426,"qr":0},"2026-03":{"total":164,"count":31,"cash":164,"qr":0}},"IG01690005":{"2026-01":{"total":491,"count":147,"cash":483,"qr":8},"2026-02":{"total":1271,"count":385,"cash":1264,"qr":7},"2026-03":{"total":304,"count":86,"cash":304,"qr":0}},"IG01690006":{"2026-01":{"total":633,"count":120,"cash":625,"qr":8},"2026-02":{"total":1592,"count":197,"cash":1372,"qr":220},"2026-03":{"total":581,"count":86,"cash":522,"qr":59}},"IG01690007":{"2026-01":{"total":113,"count":42,"cash":111,"qr":2},"2026-02":{"total":336,"count":105,"cash":320,"qr":16},"2026-03":{"total":113,"count":20,"cash":103,"qr":10}},"IG01690008":{"2026-01":{"total":55,"count":10,"cash":54,"qr":1},"2026-02":{"total":184,"count":45,"cash":184,"qr":0},"2026-03":{"total":91,"count":18,"cash":75,"qr":16}},"IG01690009":{"2026-01":{"total":61,"count":21,"cash":58,"qr":3},"2026-02":{"total":322,"count":102,"cash":322,"qr":0},"2026-03":{"total":117,"count":32,"cash":117,"qr":0}},"IG01690010":{"2026-01":{"total":110,"count":30,"cash":109,"qr":1},"2026-02":{"total":657,"count":116,"cash":657,"qr":0},"2026-03":{"total":263,"count":48,"cash":263,"qr":0}},"IG01690011":{"2026-01":{"total":436,"count":67,"cash":434,"qr":2},"2026-02":{"total":870,"count":115,"cash":870,"qr":0},"2026-03":{"total":269,"count":50,"cash":269,"qr":0}},"IG01690012":{"2026-01":{"total":316,"count":63,"cash":314,"qr":2},"2026-02":{"total":1142,"count":219,"cash":1142,"qr":0},"2026-03":{"total":261,"count":59,"cash":261,"qr":0}},"IG01690013":{"2026-01":{"total":69,"count":20,"cash":64,"qr":5},"2026-02":{"total":366,"count":71,"cash":346,"qr":20},"2026-03":{"total":248,"count":32,"cash":231,"qr":17}},"IG01690014":{"2026-01":{"total":238,"count":46,"cash":235,"qr":3},"2026-02":{"total":1044,"count":158,"cash":1025,"qr":19},"2026-03":{"total":422,"count":72,"cash":412,"qr":10}},"IG01690015":{"2026-01":{"total":198,"count":34,"cash":178,"qr":20},"2026-02":{"total":457,"count":76,"cash":457,"qr":0},"2026-03":{"total":123,"count":25,"cash":118,"qr":5}},"IG01690016":{"2026-01":{"total":38,"count":7,"cash":36,"qr":2},"2026-02":{"total":55,"count":9,"cash":55,"qr":0}},"IG01690017":{"2026-01":{"total":137,"count":36,"cash":136,"qr":1},"2026-02":{"total":377,"count":98,"cash":377,"qr":0},"2026-03":{"total":124,"count":28,"cash":124,"qr":0}},"IG01690018":{"2026-01":{"total":207,"count":45,"cash":206,"qr":1},"2026-02":{"total":824,"count":151,"cash":823,"qr":1},"2026-03":{"total":218,"count":34,"cash":218,"qr":0}},"IG01690019":{"2026-01":{"total":18,"count":2,"cash":17,"qr":1},"2026-02":{"total":291,"count":58,"cash":291,"qr":0},"2026-03":{"total":302,"count":50,"cash":302,"qr":0}},"IG01690020":{"2026-01":{"total":57,"count":14,"cash":53,"qr":4}},"IG01690021":{"2026-02":{"total":138,"count":35,"cash":137,"qr":1},"2026-03":{"total":176,"count":27,"cash":133,"qr":43}},"IG01690022":{"2026-02":{"total":330,"count":73,"cash":329,"qr":1},"2026-03":{"total":210,"count":38,"cash":210,"qr":0}},"IG01690023":{"2026-02":{"total":213,"count":27,"cash":133,"qr":80},"2026-03":{"total":30,"count":5,"cash":30,"qr":0}},"IG01690024":{"2026-02":{"total":17,"count":4,"cash":16,"qr":1}},"IG01690025":{"2026-01":{"total":25,"count":3,"cash":24,"qr":1}},"IG01690026":{"2026-02":{"total":37,"count":7,"cash":36,"qr":1},"2026-03":{"total":24,"count":7,"cash":24,"qr":0}},"IG01690027":{"2026-02":{"total":111,"count":36,"cash":90,"qr":21},"2026-03":{"total":116,"count":35,"cash":94,"qr":22}},"IG01690028":{"2026-01":{"total":51,"count":5,"cash":50,"qr":1},"2026-02":{"total":130,"count":31,"cash":130,"qr":0},"2026-03":{"total":59,"count":19,"cash":59,"qr":0}},"IG01690029":{"2026-01":{"total":37,"count":3,"cash":36,"qr":1},"2026-02":{"total":110,"count":22,"cash":110,"qr":0},"2026-03":{"total":122,"count":25,"cash":96,"qr":26}},"IG01690030":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":377,"count":67,"cash":343,"qr":34},"2026-03":{"total":185,"count":28,"cash":165,"qr":20}},"IG01690031":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":259,"count":96,"cash":228,"qr":31},"2026-03":{"total":200,"count":43,"cash":140,"qr":60}},"IG01690032":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":39,"count":10,"cash":39,"qr":0}},"IG01690033":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":116,"count":29,"cash":116,"qr":0},"2026-03":{"total":168,"count":25,"cash":138,"qr":30}},"IG01690034":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":257,"count":46,"cash":255,"qr":2},"2026-03":{"total":179,"count":27,"cash":171,"qr":8}},"IG01690035":{"2026-01":{"total":42,"count":3,"cash":41,"qr":1},"2026-02":{"total":872,"count":192,"cash":869,"qr":3},"2026-03":{"total":543,"count":134,"cash":532,"qr":11}},"IG01690036":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":182,"count":41,"cash":175,"qr":7},"2026-03":{"total":68,"count":17,"cash":61,"qr":7}},"IG01690037":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":97,"count":29,"cash":97,"qr":0},"2026-03":{"total":106,"count":24,"cash":76,"qr":30}},"IG01690038":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":407,"count":136,"cash":407,"qr":0},"2026-03":{"total":325,"count":79,"cash":315,"qr":10}},"IG01690039":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":194,"count":48,"cash":194,"qr":0},"2026-03":{"total":299,"count":63,"cash":299,"qr":0}},"IG01690040":{"2026-01":{"total":19,"count":2,"cash":18,"qr":1},"2026-02":{"total":38,"count":11,"cash":38,"qr":0},"2026-03":{"total":151,"count":28,"cash":141,"qr":10}}};

// ── Helpers ───────────────────────────────────────────────────────────────────
const THAI_M = ["","ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
const MCOLS  = ["#3b82f6","#0d9488","#f59e0b","#a78bfa","#fb923c","#34d399"];
const BG     = "#ffffff";
const CARD   = "#ffffff";
const BORDER = "#d8e0e8";

const ml = m => { const [y,mo]=m.split("-"); return `${THAI_M[+mo]} ${String(+y+543).slice(-2)}`; };
const fmt = n => Math.round(n||0).toLocaleString("th-TH");
const getMonths = data => [...new Set(Object.values(data).flatMap(d=>Object.keys(d)))].sort();

function Trend({ t, v }) {
  if (!t||t==="none"||!v) return null;
  if (t==="up")   return <span style={{color:"#0d9488",fontSize:11}}>▲</span>;
  if (t==="down") return <span style={{color:"#ef4444",fontSize:11}}>▼</span>;
  if (t==="new")  return <span style={{color:"#60a5fa",fontSize:9,fontWeight:700}}>NEW</span>;
  return null;
}

// Parse CSV text → same structure as INIT_DATA
function parseCSV(text) {
  const lines = text.trim().split("\n");
  // find header
  let hi = -1;
  for (let i=0;i<Math.min(15,lines.length);i++) {
    const l = lines[i].toLowerCase();
    if (l.includes("device") && l.includes("amount")) { hi=i; break; }
  }
  if (hi<0) throw new Error("ไม่พบ header row");
  const headers = lines[hi].split(",").map(h=>h.replace(/"/g,"").trim().toLowerCase());
  const cDev = headers.findIndex(h=>h.includes("device detail"));
  const cAmt = headers.findIndex(h=>h==="amount");
  const cTyp = headers.findIndex(h=>h==="type");
  const cSts = headers.findIndex(h=>h==="status");
  const cDat = headers.findIndex(h=>h==="date");
  if (cDev<0||cAmt<0) throw new Error("ไม่พบคอลัมน์ Device Detail หรือ Amount");
  const result = {};
  for (let i=hi+1;i<lines.length;i++) {
    const row = lines[i].split(",").map(c=>c.replace(/"/g,"").trim());
    if (!row[cSts]||row[cSts].toLowerCase()!=="success") continue;
    const dev = row[cDev]; if (!dev) continue;
    const amt = parseFloat(row[cAmt])||0;
    const typ = (row[cTyp]||"").toLowerCase();
    const ds = row[cDat]||"";
    const mm = ds.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (!mm) continue;
    const key = `${mm[3]}-${mm[2].padStart(2,"0")}`;
    if (!result[dev]) result[dev]={};
    if (!result[dev][key]) result[dev][key]={total:0,count:0,cash:0,qr:0};
    result[dev][key].total += amt;
    result[dev][key].count += 1;
    if (typ==="cash") result[dev][key].cash += amt;
    else if (typ==="qr") result[dev][key].qr += amt;
  }
  Object.values(result).forEach(d=>Object.values(d).forEach(v=>{
    v.total=Math.round(v.total*10)/10; v.cash=Math.round(v.cash*10)/10; v.qr=Math.round(v.qr*10)/10;
  }));
  return result;
}

// ── MapModal: OpenStreetMap embed ────────────────────────────────────────────
function MapModal({ item, onClose }) {
  if (!item) return null;
  const q = encodeURIComponent((item.label || "") + " ประเทศไทย");
  const osmSrc = `https://www.openstreetmap.org/export/embed.html?layer=mapnik&query=${q}`;
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.8)",zIndex:99999,display:"flex",alignItems:"center",justifyContent:"center",padding:12}}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#ffffff",border:"1px solid #d8e0e8",borderRadius:14,width:"min(860px,96vw)",height:"80vh",display:"flex",flexDirection:"column",boxShadow:"0 8px 48px rgba(0,0,0,0.7)"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"11px 16px",borderBottom:"1px solid #d8e0e8",flexShrink:0}}>
          <span style={{fontWeight:700,color:"#0f1824",fontSize:14}}>📍 {item.label}</span>
          <button onClick={onClose} style={{background:"rgba(15,24,36,0.04)",border:"1px solid #d8e0e8",color:"#5b7186",fontSize:16,cursor:"pointer",borderRadius:7,width:30,height:30,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
        </div>
        <iframe src={osmSrc} style={{flex:1,border:"none",borderRadius:"0 0 14px 14px"}} allowFullScreen loading="lazy"/>
      </div>
    </div>
  );
}


// ── PhotoCell: open Drive folder directly ────────────────────────────────────
function PhotoCell({ url, label }) {
  const [show, setShow] = useState(false);
  return (
    <div style={{position:"relative",display:"inline-flex",alignItems:"center",justifyContent:"center"}}
      onMouseEnter={()=>setShow(true)} onMouseLeave={()=>setShow(false)}>
      <a href={url} target="_blank" rel="noopener noreferrer"
        style={{display:"inline-flex",alignItems:"center",justifyContent:"center",
          width:30,height:30,background:show?"rgba(45,212,191,0.1)":"#f1f5f9",
          border:`1px solid ${show?"#0d9488":"#d8e0e8"}`,
          borderRadius:8,fontSize:16,textDecoration:"none",transition:"all .15s"}}>
        🖼️
      </a>
      {show&&<div style={{position:"absolute",bottom:"calc(100% + 8px)",left:"50%",transform:"translateX(-50%)",
        background:"#f1f5f9",color:"#1e2a3a",fontSize:11,padding:"5px 10px",borderRadius:8,
        whiteSpace:"nowrap",maxWidth:220,textAlign:"center",zIndex:9999,pointerEvents:"none",
        boxShadow:"0 4px 16px rgba(0,0,0,0.3)",lineHeight:1.5}}>
        {label}
        <div style={{position:"absolute",top:"100%",left:"50%",transform:"translateX(-50%)",
          borderLeft:"5px solid transparent",borderRight:"5px solid transparent",borderTop:"5px solid #f1f5f9"}}/>
      </div>}
    </div>
  );
}


// ── Export ranking data to Excel ─────────────────────────────────────────────
async function exportRankingExcel(rankData, months, rentMap, locMap, install, elapsed, lastDays, ml) {
  // SheetJS is bundled with the app (imported at module scope)
  const lastM = months[months.length - 1];
  const isLastOpen = elapsed < lastDays;

  // Build rows
  const headers = [
    "#", "หมายเลขเครื่อง", "สถานที่", "วันติดตั้ง", "เดือน(ติดตั้ง)", "ค่าเช่า",
    ...months.map((m, i) => ml(m) + (i === months.length - 1 && isLastOpen ? " *" : "")),
    "AVG", "P/L", "รวม", "ครั้ง", "สัดส่วน%", "สถานะ"
  ];

  const totalGrand = rankData.reduce((s, r) => s + r.total, 0);

  const rows = rankData.map((r, i) => {
    const rent = rentMap[r.device] || 0;
    const isLow = r.mature && rent > 0 && r.avg < rent;
    const monthVals = months.map((m, mi) => {
      const isLast = mi === months.length - 1 && isLastOpen;
      return isLast ? r.rr : (r.bm[m] || 0);
    });
    return [
      i + 1,
      r.device,
      locMap[r.device] || "",
      install[r.device] || "",
      r.monthsInstalled,
      rent,
      ...monthVals,
      r.avg,
      Math.round(r.avg - rent),
      r.total,
      r.count,
      parseFloat((r.total / totalGrand * 100).toFixed(1)),
      isLow ? "⚠️ ต่ำกว่าค่าเช่า" : r.mature ? "✅ ปกติ" : `ใหม่ (${r.monthsInstalled}M)`
    ];
  });

  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

  // Column widths
  ws["!cols"] = [
    {wch:4},{wch:14},{wch:28},{wch:12},{wch:8},{wch:10},
    ...months.map(()=>({wch:12})),
    {wch:10},{wch:10},{wch:10},{wch:8},{wch:10},{wch:16}
  ];

  // Style header row (bold)
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let C = range.s.c; C <= range.e.c; C++) {
    const addr = XLSX.utils.encode_cell({r:0, c:C});
    if (!ws[addr]) continue;
    ws[addr].s = { font:{bold:true}, fill:{fgColor:{rgb:"1E2A3A"}}, alignment:{horizontal:"center"} };
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "จัดอันดับรายได้");

  // Add summary sheet
  const summaryData = [
    ["สรุปข้อมูล", ""],
    ["จำนวนเครื่องทั้งหมด", rankData.length],
    ["รายได้รวม (บาท)", rankData.reduce((s,r)=>s+r.total, 0)],
    ["AVG เฉลี่ยทุกเครื่อง", Math.round(rankData.reduce((s,r)=>s+r.avg,0)/rankData.length)],
    ["", ""],
    ["เครื่องที่ AVG < ค่าเช่า", rankData.filter(r=>r.mature&&(rentMap[r.device]||0)>0&&r.avg<(rentMap[r.device]||0)).length],
    ["เครื่องที่ติดตั้ง ≥ 2 เดือน", rankData.filter(r=>r.mature).length],
    ["เครื่องที่ยังใหม่ < 2 เดือน", rankData.filter(r=>!r.mature).length],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(summaryData);
  ws2["!cols"] = [{wch:28},{wch:16}];
  XLSX.utils.book_append_sheet(wb, ws2, "สรุป");

  const now = new Date();
  const dateStr = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,"0")}${String(now.getDate()).padStart(2,"0")}`;

  // Build the workbook as a Blob and trigger a download via an anchor.
  // XLSX.writeFile() does an in-frame <a download> click, which sandboxed
  // hosts (e.g. the Claude artifact iframe) silently block. When we detect we
  // are inside an iframe, open the file in a new top-level tab instead — the
  // same target="_blank" path the photo links already use successfully.
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const inIframe = (() => { try { return window.self !== window.top; } catch (e) { return true; } })();
  const a = document.createElement("a");
  a.href = url;
  a.download = `ranking_${dateStr}.xlsx`;
  if (inIframe) { a.target = "_blank"; a.rel = "noopener"; }
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

// ── Main Component ────────────────────────────────────────────────────────────
export default function App() {
  const [data,     setData]    = useState(() => LS.get("bt_data", {}));
  const [fname,    setFname]   = useState(() => LS.get("bt_fname", ""));
  const [install,  setInstall] = useState(() => LS.get("bt_install", DEFAULT_INSTALL));
  // Start from the built-in reference data, then overlay any saved/uploaded data
  // so location, map and photo columns show up even before an area file is uploaded.
  const [mapUrl,   setMapUrl]   = useState(() => ({ ...DEFAULT_MAP,      ...LS.get("bt_mapUrl", {}) }));
  const [locMap,   setLocMap]   = useState(() => ({ ...DEFAULT_LOCATION, ...LS.get("bt_locMap", {}) }));
  const [photoMap, setPhotoMap] = useState(() => ({ ...DEFAULT_PHOTO,    ...LS.get("bt_photoMap", {}) }));
  const [folderEmbed, setFolderEmbed] = useState(() => LS.get("bt_folderEmbed", DEFAULT_FOLDER_EMBED));
  const [rentMap,  setRentMap] = useState(() => LS.get("bt_rentMap", DEFAULT_RENT));
  const [mfname,   setMfname]  = useState(() => LS.get("bt_mfname", "ข้อมูลตัวอย่าง"));
  const [view,     setView]    = useState("overview");
  const [selDev,   setSelDev]  = useState("");
  const [rankSort,   setRankSort]  = useState({col:"pl",dir:"desc"});
  const [exporting,  setExporting] = useState(false);
  const [copied,   setCopied]   = useState("");
  const [mapModal, setMapModal] = useState(null); // {url, label}
  const [drag,     setDrag]    = useState(false);
  const [dragM,    setDragM]   = useState(false);
  const [busy,     setBusy]    = useState(false);
  const [busyM,    setBusyM]   = useState(false);
  const [err,      setErr]     = useState("");
  const [errM,     setErrM]    = useState("");
  const fileRef  = useRef();
  const fileRefM = useRef();

  // Persist the sales-data channel whenever it changes.
  useEffect(() => { LS.set("bt_data", data); LS.set("bt_fname", fname); }, [data, fname]);
  // Persist the area / machine-info channel whenever it changes.
  useEffect(() => {
    LS.set("bt_install", install);
    LS.set("bt_rentMap", rentMap);
    LS.set("bt_mapUrl", mapUrl);
    LS.set("bt_locMap", locMap);
    LS.set("bt_photoMap", photoMap);
    LS.set("bt_folderEmbed", folderEmbed);
    LS.set("bt_mfname", mfname);
  }, [install, rentMap, mapUrl, locMap, photoMap, folderEmbed, mfname]);

  // ── Cloud sync (Firebase) ─────────────────────────────────────────────────
  // A live mirror of the current state so cloud writes always use fresh values.
  const stateRef = useRef({});
  useEffect(() => {
    stateRef.current = { data, fname, install, rentMap, mapUrl, locMap, photoMap, folderEmbed, mfname };
  });
  const [cloudStatus, setCloudStatus] = useState(cloudEnabled ? "connecting" : "off");
  const cloudReadyRef = useRef(false);
  // Write the full report (current state merged with overrides) to the cloud.
  const pushCloud = useCallback((overrides = {}) => {
    if (!cloudEnabled) return;
    const payload = { ...stateRef.current, ...overrides };
    saveReport(payload).then(() => setCloudStatus("synced")).catch(() => setCloudStatus("error"));
  }, []);
  // Subscribe to the shared cloud report and apply remote changes to local state.
  useEffect(() => {
    if (!cloudEnabled) return;
    const unsub = subscribeReport((p) => {
      if (p) {
        if (p.data) setData(p.data);
        if (p.fname != null) setFname(p.fname);
        if (p.install) setInstall(p.install);
        if (p.rentMap) setRentMap(p.rentMap);
        if (p.mapUrl) setMapUrl(p.mapUrl);
        if (p.locMap) setLocMap(p.locMap);
        if (p.photoMap) setPhotoMap(p.photoMap);
        if (p.folderEmbed) setFolderEmbed(p.folderEmbed);
        if (p.mfname != null) setMfname(p.mfname);
        setCloudStatus("synced");
      } else {
        // No cloud document yet — seed it from whatever we have locally.
        setCloudStatus("synced");
        if (stateRef.current.data && Object.keys(stateRef.current.data).length) pushCloud();
      }
      cloudReadyRef.current = true;
    }, () => setCloudStatus("error"));
    return () => unsub();
  }, [pushCloud]);

  const months  = useMemo(()=>getMonths(data),[data]);
  const devices = useMemo(()=>Object.keys(data).sort(),[data]);

  // On load, if we restored sales data but have no selected device yet, pick the first.
  useEffect(() => {
    if (!selDev && devices.length) setSelDev(devices[0]);
  }, [devices, selDev]);
  const lastM   = months[months.length-1];
  const lastDays = useMemo(()=>{ if(!lastM)return 31; const[y,mo]=lastM.split("-").map(Number); return new Date(y,mo,0).getDate(); },[lastM]);
  const elapsed  = useMemo(()=>{ if(!lastM)return lastDays; const[y,mo]=lastM.split("-").map(Number); const t=new Date(); return t.getFullYear()===y&&t.getMonth()+1===mo?t.getDate():lastDays; },[lastM,lastDays]);

  // ── Executive summary (buckets by monthly revenue) for LINE ─────────────────
  const [sumMonth,  setSumMonth]  = useState("");
  const [copiedSum, setCopiedSum] = useState(false);
  // Default the summary month to the last *complete* month (skip the partial current month).
  useEffect(() => {
    if (!months.length) return;
    const lastComplete = (elapsed < lastDays && months.length > 1) ? months[months.length-2] : months[months.length-1];
    if (!sumMonth || !months.includes(sumMonth)) setSumMonth(lastComplete);
  }, [months, elapsed, lastDays, sumMonth]);
  const summaryText = useMemo(() => {
    if (!sumMonth) return "";
    // Bucket by net revenue after rent (ยอดขาย - ค่าเช่า) for the selected month.
    const b = { over6:0, b36:0, b23:0, b12:0, b01:0, neg:0 };
    let total = 0, grossSum = 0, netSum = 0;
    devices.forEach(d => {
      const v = data[d]?.[sumMonth]?.total || 0;
      if (v <= 0) return;              // count only machines that reported that month
      const net = v - (rentMap[d] || 0);
      total++; grossSum += v; netSum += net;
      if (net >= 6000) b.over6++;
      else if (net >= 3000) b.b36++;
      else if (net >= 2000) b.b23++;
      else if (net >= 1000) b.b12++;
      else if (net >= 0) b.b01++;
      else b.neg++;
    });
    const lines = [
      `📊 สรุปตู้ Bottalk — เดือน ${ml(sumMonth)}`,
      `(สุทธิหักเช่า = ยอดขาย - ค่าเช่า)`,
      ``,
      `จำนวนตู้ทั้งหมด = ${total} ตู้`,
    ];
    if (b.over6) lines.push(`ตู้สุทธิ มากกว่า 6,000 = ${b.over6} ตู้`);
    lines.push(`ตู้สุทธิ 3,000-6,000 = ${b.b36} ตู้`);
    lines.push(`ตู้สุทธิ 2,000-2,999 = ${b.b23} ตู้`);
    lines.push(`ตู้สุทธิ 1,000-1,999 = ${b.b12} ตู้`);
    lines.push(`ตู้สุทธิ ต่ำกว่า 1,000 = ${b.b01} ตู้`);
    lines.push(`ตู้สุทธิ ติดลบ (ขาดทุน) = ${b.neg} ตู้`);
    lines.push(``, `รายได้รวม (ยอดขาย) = ฿${fmt(grossSum)}`, `สุทธิหักค่าเช่า = ฿${fmt(netSum)}`);
    return lines.join("\n");
  }, [sumMonth, devices, data, rentMap]);
  const copySummary = () => {
    const text = summaryText;
    try {
      const el = document.createElement("textarea");
      el.value = text; el.style.position = "fixed"; el.style.opacity = "0";
      document.body.appendChild(el); el.focus(); el.select();
      document.execCommand("copy");
      document.body.removeChild(el);
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).catch(()=>{});
    } catch (e) {}
    setCopiedSum(true); setTimeout(() => setCopiedSum(false), 1800);
  };

  // Load file — supports .xlsx via SheetJS CDN (loaded lazily), .csv natively
  const loadFile = useCallback(async file => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    setBusy(true); setErr("");
    try {
      if (ext==="csv") {
        const text = await file.text();
        const parsed = parseCSV(text);
        setData(parsed); setFname(file.name); setSelDev(Object.keys(parsed).sort()[0]); setView("overview");
        pushCloud({ data: parsed, fname: file.name });
      } else if (ext==="xlsx"||ext==="xls") {
        const buf = await file.arrayBuffer();
        const wb  = XLSX.read(buf,{type:"array"});
        const ws  = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
        let hi=-1;
        for(let i=0;i<Math.min(15,rows.length);i++){
          const r=rows[i].map(c=>String(c).toLowerCase());
          if(r.some(c=>c.includes("device"))&&r.some(c=>c.includes("amount"))){hi=i;break;}
        }
        if(hi<0) throw new Error("ไม่พบ header row");
        const hdrs=rows[hi].map(c=>String(c).trim().toLowerCase());
        const cDev=hdrs.findIndex(h=>h.includes("device detail"));
        const cAmt=hdrs.findIndex(h=>h==="amount");
        const cTyp=hdrs.findIndex(h=>h==="type");
        const cSts=hdrs.findIndex(h=>h==="status");
        const cDat=hdrs.findIndex(h=>h==="date");
        if(cDev<0||cAmt<0) throw new Error("ไม่พบคอลัมน์ Device Detail หรือ Amount");
        const result={};
        for(let i=hi+1;i<rows.length;i++){
          const row=rows[i];
          if(String(row[cSts]||"").trim().toLowerCase()!=="success") continue;
          const dev=String(row[cDev]||"").trim(); if(!dev) continue;
          const amt=parseFloat(row[cAmt])||0;
          const typ=String(row[cTyp]||"").trim().toLowerCase();
          const ds=String(row[cDat]||"").trim();
          const mm=ds.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
          if(!mm) continue;
          const key=`${mm[3]}-${mm[2].padStart(2,"0")}`;
          if(!result[dev]) result[dev]={};
          if(!result[dev][key]) result[dev][key]={total:0,count:0,cash:0,qr:0};
          result[dev][key].total+=amt; result[dev][key].count+=1;
          if(typ==="cash") result[dev][key].cash+=amt;
          else if(typ==="qr") result[dev][key].qr+=amt;
        }
        Object.values(result).forEach(d=>Object.values(d).forEach(v=>{
          v.total=Math.round(v.total*10)/10; v.cash=Math.round(v.cash*10)/10; v.qr=Math.round(v.qr*10)/10;
        }));
        setData(result); setFname(file.name); setSelDev(Object.keys(result).sort()[0]); setView("overview");
        pushCloud({ data: result, fname: file.name });
      } else {
        throw new Error("รองรับเฉพาะ .xlsx, .xls, .csv");
      }
    } catch(e) { setErr(e.message); }
    setBusy(false);
  },[pushCloud]);

  // Load machine-info file (install date + rent)
  const loadMachineFile = useCallback(async file => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    setBusyM(true); setErrM("");
    try {
      let rows = [];
      if (ext === "csv") {
        const text = await file.text();
        rows = text.trim().split("\n").map(l => l.split(",").map(c => c.replace(/"/g,"").trim()));
      } else if (ext === "xlsx" || ext === "xls") {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, {type:"array", cellDates:false});
        const ws = wb.Sheets[wb.SheetNames[0]];
        rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:""});
      } else {
        throw new Error("รองรับเฉพาะ .xlsx, .xls, .csv");
      }
      const { install: newInstall, rent: newRent, mapData: newMapData, locData: newLocData, photoData: newPhotoData } = parseMachineInfo(rows);
      if (Object.keys(newInstall).length === 0 && Object.keys(newRent).length === 0)
        throw new Error("ไม่พบข้อมูลในไฟล์ — ตรวจสอบ column headers");
      // Replace each channel the new file provides (new data fully replaces old);
      // keep the existing values only for channels the file doesn't include.
      const cur = stateRef.current;
      const mInstall = Object.keys(newInstall).length ? newInstall : cur.install;
      const mRent    = Object.keys(newRent).length    ? newRent    : cur.rentMap;
      const mMap     = (newMapData   && Object.keys(newMapData).length)   ? newMapData   : cur.mapUrl;
      const mLoc     = (newLocData   && Object.keys(newLocData).length)   ? newLocData   : cur.locMap;
      const mPhoto   = (newPhotoData && Object.keys(newPhotoData).length) ? newPhotoData : cur.photoMap;
      setInstall(mInstall);
      setRentMap(mRent);
      setMapUrl(mMap);
      setLocMap(mLoc);
      setPhotoMap(mPhoto);
      setMfname(file.name);
      pushCloud({ install: mInstall, rentMap: mRent, mapUrl: mMap, locMap: mLoc, photoMap: mPhoto, mfname: file.name });
    } catch(e) { setErrM(e.message); }
    setBusyM(false);
  }, [pushCloud]);

  // ── Computed ─────────────────────────────────────────────────────────────────
  const ovData = useMemo(()=>months.map(m=>{
    let total=0,cash=0,qr=0,count=0;
    devices.forEach(d=>{const r=data[d]?.[m];if(r){total+=r.total;cash+=r.cash;qr+=r.qr;count+=r.count;}});
    return {month:ml(m),total,cash,qr,count};
  }),[data,months,devices]);

  const devData = useMemo(()=>months.map(m=>{
    const r=data[selDev]?.[m]||{total:0,cash:0,qr:0,count:0};
    return {month:ml(m),...r};
  }),[data,months,selDev]);

  const rankData = useMemo(()=>devices.map(d=>{
    let total=0,count=0; const bm={};
    months.forEach(m=>{const r=data[d]?.[m]; bm[m]=r?r.total:0; if(r){total+=r.total;count+=r.count;}});
    const act=months.filter(m=>bm[m]>0).length;
    const rr=bm[lastM]>0?Math.round(bm[lastM]/elapsed*lastDays):0;
    // AVG: for last month use RunRate projection if month not complete
    const lastMonthIsOpen = elapsed < lastDays;
    const totalForAvg = months.reduce((s,m)=>{
      if(m===lastM && lastMonthIsOpen) return s + (rr||bm[lastM]||0);
      return s + (bm[m]||0);
    }, 0);
    const avg=act>0?Math.round(totalForAvg/act):0;
    const tr={};
    for(let i=1;i<months.length;i++){
      const p=bm[months[i-1]],c=bm[months[i]];
      tr[months[i]]=c===0?"none":p===0?"new":c>p?"up":c<p?"down":"flat";
    }
    // คำนวณจำนวนเดือนที่ติดตั้ง (นับจาก install date ถึงเดือนล่าสุดในข้อมูล)
    const instStr = install[d];
    let monthsInstalled = 0;
    if (instStr && lastM) {
      const [dd,mm,yy] = instStr.split("-").map(Number);
      const instDate = new Date(2000+yy, mm-1, dd);
      const [lY,lM] = lastM.split("-").map(Number);
      const lastDate = new Date(lY, lM-1, 1);
      monthsInstalled = (lastDate.getFullYear()-instDate.getFullYear())*12
                      + (lastDate.getMonth()-instDate.getMonth());
    }
    const mature = monthsInstalled >= 2; // ติดตั้งมาแล้ว >= 2 เดือน
    return {device:d,total,count,bm,avg,rr,tr,monthsInstalled,mature};
  }).sort((a,b)=>{
    if(a.mature!==b.mature) return a.mature?-1:1;
    const dir = rankSort.dir==="asc"?1:-1;
    const col = rankSort.col;
    if(col==="avg")   return dir*(a.avg-b.avg);
    if(col==="pl")    return dir*((a.avg-(rentMap[a.device]||0))-(b.avg-(rentMap[b.device]||0)));
    if(col==="total") return dir*(a.total-b.total);
    if(col==="rr")    return dir*(a.rr-b.rr);
    if(col==="count") return dir*(a.count-b.count);
    if(col==="rent")  return dir*((rentMap[a.device]||0)-(rentMap[b.device]||0));
    if(col==="loc")   return dir*((locMap[a.device]||"").localeCompare(locMap[b.device]||"","th"));
    if(col==="device")return dir*(a.device.localeCompare(b.device));
    return dir*(a.avg-b.avg);
  }),[data,months,devices,lastM,elapsed,lastDays,install,rankSort,rentMap,locMap]);

  const grandTotal = useMemo(()=>rankData.reduce((s,r)=>s+r.total,0),[rankData]);

  // Count machines by P/L tier (AVG - rent) for the ranking overview boxes.
  const plTiers = useMemo(()=>{
    const t = { loss:0, t0:0, t1:0, t2:0, t3:0 };
    rankData.forEach(r=>{
      const pl = Math.round(r.avg - (rentMap[r.device]||0));
      if (pl < 0) t.loss++;
      else if (pl <= 1000) t.t0++;
      else if (pl <= 2000) t.t1++;
      else if (pl <= 3000) t.t2++;
      else t.t3++;
    });
    return t;
  },[rankData,rentMap]);

  const stackData = useMemo(()=>rankData.map(r=>{
    const o={device:r.device}; months.forEach(m=>{o[ml(m)]=r.bm[m]||0;}); return o;
  }),[rankData,months]);

  const pieData = useMemo(()=>{
    let cash=0,qr=0; devData.forEach(r=>{cash+=r.cash;qr+=r.qr;});
    return [{name:"เงินสด",value:cash},{name:"QR",value:qr}];
  },[devData]);

  const decData = useMemo(()=>{
    if(!rankData.length) return [];
    return rankData.map(r=>{
    const rent=rentMap[r.device]||0;
    const avail=months.filter(m=>r.bm[m]>0);
    const miss=avail.length>0&&!r.bm[lastM];
    const cd=avail.length>=3&&r.bm[avail.at(-1)]<r.bm[avail.at(-2)]&&r.bm[avail.at(-2)]<r.bm[avail.at(-3)];
    let status,reason;
    if(miss)                          {status="nodata";  reason="ไม่มีข้อมูลเดือนล่าสุด — ตรวจสอบสถานะ";}
    else if(r.rr>=rent&&r.avg>=rent)  {status="keep";    reason="Run Rate และ AVG ผ่านค่าเช่า";}
    else if(r.rr>=rent||r.avg>=rent)  {status="watch";   reason=r.rr>=rent?"Run Rate ผ่าน แต่ AVG ยังต่ำกว่าค่าเช่า":"AVG ผ่าน แต่ Run Rate เดือนนี้ต่ำ";}
    else if(cd)                       {status="relocate"; reason="ยอดลดติดต่อกัน 3 เดือนและต่ำกว่าค่าเช่า";}
    else                              {status="relocate"; reason="AVG และ Run Rate ต่ำกว่าค่าเช่า";}
    return {...r,rent,status,reason};
  });
  },[rankData,lastM,months,rentMap]);

  const grp = useMemo(()=>({
    keep:decData.filter(r=>r.status==="keep"),
    watch:decData.filter(r=>r.status==="watch"),
    relocate:decData.filter(r=>r.status==="relocate"),
    nodata:decData.filter(r=>r.status==="nodata"),
  }),[decData]);

  // ── Styles ────────────────────────────────────────────────────────────────────
  const tab = t => ({
    padding:"7px 16px",borderRadius:8,border:"none",cursor:"pointer",
    fontFamily:"inherit",fontWeight:600,fontSize:13,transition:"all .15s",
    background:view===t?"#0d9488":"#e6ebf1",
    color:view===t?"#fff":"#64748b",
  });
  const cv = (r,m,showT,isLastOpen) => {
    const actual = r.bm[m]||0;
    const v = isLastOpen ? r.rr : actual;
    const t = r.tr[m];
    return <td key={m} style={{padding:"8px 9px",textAlign:"right",color:v>0?(isLastOpen?"#2563eb":"#5b7186"):"#d8e0e8",fontWeight:isLastOpen&&v>0?700:400,whiteSpace:"nowrap"}}>
      <span style={{display:"inline-flex",alignItems:"center",gap:3,justifyContent:"flex-end"}}>
        {v>0?`฿${fmt(v)}`:"—"}{showT&&!isLastOpen&&<Trend t={t} v={actual}/>}
      </span>
    </td>;
  };

  // ── Render ────────────────────────────────────────────────────────────────────
  return (
    <div style={{minHeight:"100vh",background:BG,color:"#1e2a3a",fontFamily:"'IBM Plex Sans Thai','Sarabun',sans-serif",padding:"18px 16px"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&display=swap');
        html,body{background:${BG}!important;margin:0}
        *{box-sizing:border-box}
        ::-webkit-scrollbar{width:5px;height:5px}
        ::-webkit-scrollbar-track{background:#e2e8f0}
        ::-webkit-scrollbar-thumb{background:#d8e0e8;border-radius:3px}
        .card{background:${CARD};border:1px solid ${BORDER};border-radius:14px;padding:16px;box-shadow:0 1px 3px rgba(15,24,36,0.08)}
        .sc{border-radius:12px;padding:14px 18px}
        select{background:#ffffff;color:#1e2a3a;border:1px solid #d8e0e8;border-radius:9px;padding:6px 12px;font-size:13px;font-family:inherit;outline:none;cursor:pointer;box-shadow:0 1px 2px rgba(15,24,36,0.08)}
        .rr{transition:background .12s;cursor:pointer;border-bottom:1px solid #e6ebf1}
        .rr:hover td{background:rgba(13,148,136,0.06)}
        .rr-low{background:rgba(220,38,38,0.08)}
        .rr-low:hover td{background:rgba(220,38,38,0.13)!important}
        .rr-sel{background:rgba(13,148,136,0.1);box-shadow:inset 3px 0 0 #0d9488}
        .dz{border:2px dashed #d8e0e8;border-radius:10px;padding:7px 13px;cursor:pointer;transition:all .18s;background:#ffffff;display:inline-flex;align-items:center;gap:8px}
        .dz:hover,.dz.drag{border-color:#0d9488;background:rgba(45,212,191,0.07)}
        .tip-wrap{position:relative;display:inline-flex;align-items:center;justify-content:center}
        /* Clearer table grid so rows and columns are easy to follow */
        table{border-collapse:collapse}
        th,td{border:1px solid #dbe3ec}
        thead th{border-bottom:2px solid #aebccb;background:#f2f6fa}
        tbody tr:hover td{background:#eef4f9}
      `}</style>

      {/* Header */}
      <div style={{marginBottom:18,display:"flex",alignItems:"flex-start",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <h1 style={{margin:"0 0 3px",fontSize:21,fontWeight:700,color:"#0f1824"}}>วิเคราะห์รายได้ตู้กดน้ำ</h1>
            {cloudStatus!=="off"&&(
              <span title="ข้อมูลถูกบันทึกบนคลาวด์ (ทุกเครื่องเห็นเหมือนกัน)" style={{fontSize:10.5,fontWeight:600,padding:"2px 8px",borderRadius:20,whiteSpace:"nowrap",
                background:cloudStatus==="synced"?"rgba(45,212,191,0.12)":cloudStatus==="error"?"rgba(220,38,38,0.1)":"rgba(100,116,139,0.1)",
                color:cloudStatus==="synced"?"#0d9488":cloudStatus==="error"?"#ef4444":"#64748b",
                border:`1px solid ${cloudStatus==="synced"?"rgba(45,212,191,0.3)":cloudStatus==="error"?"rgba(220,38,38,0.25)":"rgba(100,116,139,0.2)"}`}}>
                {cloudStatus==="synced"?"☁️ ซิงค์แล้ว":cloudStatus==="error"?"⚠️ ออฟไลน์":"☁️ กำลังเชื่อม…"}
              </span>
            )}
          </div>
          <p style={{margin:0,color:"#64748b",fontSize:11}}>{fname?`📊 ${fname}`:"📊 ยังไม่ได้อัปโหลดข้อมูล"} {mfname?`· 🏭 ${mfname}`:""} {months.length>0?`· ${devices.length} เครื่อง · ${ml(months[0])} – ${ml(lastM)}`:""}</p>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
          {[["overview","ภาพรวม"],["detail","รายเครื่อง"],["ranking","จัดอันดับ"],["decision","สรุปสถานะ"]].map(([t,l])=>(
            <button key={t} style={tab(t)} onClick={()=>setView(t)}>{l}</button>
          ))}
          {/* File 1: transaction data */}
          <div
            className={`dz${drag?" drag":""}`}
            onDragOver={e=>{e.preventDefault();setDrag(true);}}
            onDragLeave={()=>setDrag(false)}
            onDrop={e=>{e.preventDefault();setDrag(false);loadFile(e.dataTransfer.files[0]);}}
            onClick={()=>fileRef.current.click()}
          >
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>loadFile(e.target.files[0])}/>
            <span style={{fontSize:14}}>{busy?"⏳":"📊"}</span>
            <span style={{fontSize:12,color:"#64748b"}}>{busy?"กำลังโหลด…":"ข้อมูลยอดขาย"}</span>
          </div>
          {/* File 2: machine info (install date + rent) */}
          <div
            className={`dz${dragM?" drag":""}`}
            onDragOver={e=>{e.preventDefault();setDragM(true);}}
            onDragLeave={()=>setDragM(false)}
            onDrop={e=>{e.preventDefault();setDragM(false);loadMachineFile(e.dataTransfer.files[0]);}}
            onClick={()=>fileRefM.current.click()}
            style={{borderColor: mfname!=="ข้อมูลตัวอย่าง"?"#0d9488":undefined}}
          >
            <input ref={fileRefM} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>loadMachineFile(e.target.files[0])}/>
            <span style={{fontSize:14}}>{busyM?"⏳":"🏭"}</span>
            <span style={{fontSize:12,color: mfname!=="ข้อมูลตัวอย่าง"?"#0d9488":"#64748b"}}>{busyM?"กำลังโหลด…":mfname!=="ข้อมูลตัวอย่าง"?`✓ ${mfname.slice(0,18)}…`:"ข้อมูลเครื่อง"}</span>
          </div>
        </div>
      </div>

      {err&&<div style={{background:"rgba(220,38,38,0.1)",border:"1px solid rgba(220,38,38,0.3)",borderRadius:9,padding:"9px 14px",color:"#f87171",fontSize:12,marginBottom:10}}>⚠️ {err} — หากอัปโหลด .xlsx ไม่ได้ ลอง Export เป็น .csv แล้วลองใหม่</div>}
      {errM&&<div style={{background:"rgba(251,191,36,0.08)",border:"1px solid rgba(251,191,36,0.25)",borderRadius:9,padding:"9px 14px",color:"#fbbf24",fontSize:12,marginBottom:10}}>⚠️ ไฟล์ข้อมูลเครื่อง: {errM}</div>}

      {/* ── OVERVIEW ── */}
      {view==="overview"&&<>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(155px,1fr))",gap:10,marginBottom:16}}>
          <div className="sc" style={{background:"rgba(45,212,191,0.08)",border:"1px solid rgba(45,212,191,0.25)"}}>
            <div style={{fontSize:11,color:"#0d9488",fontWeight:600,marginBottom:4}}>รายได้รวมทั้งหมด</div>
            <div style={{fontSize:24,fontWeight:700,color:"#0f1824"}}>฿{fmt(grandTotal)}</div>
          </div>
          {ovData.map((m,i)=>(
            <div key={m.month} className="sc" style={{background:"rgba(99,179,237,0.08)",border:"1px solid rgba(99,179,237,0.2)"}}>
              <div style={{fontSize:11,color:MCOLS[i%MCOLS.length],fontWeight:600,marginBottom:4}}>{m.month}</div>
              <div style={{fontSize:20,fontWeight:700,color:"#0f1824"}}>฿{fmt(m.total)}</div>
              <div style={{fontSize:11,color:"#64748b",marginTop:2}}>{fmt(m.count)} ครั้ง</div>
            </div>
          ))}
        </div>

        <div className="card" style={{marginBottom:14}}>
          <div style={{fontSize:13,color:"#5b7186",fontWeight:600,marginBottom:12}}>รายได้รวมทุกเครื่อง แยกตามเดือน</div>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={ovData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#ffffff"/>
              <XAxis dataKey="month" tick={{fill:"#64748b",fontSize:12}}/>
              <YAxis tick={{fill:"#64748b",fontSize:11}} tickFormatter={v=>`฿${(v/1000).toFixed(0)}k`}/>
              <Tooltip contentStyle={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:9,color:"#1e2a3a"}} formatter={(v,n)=>[`฿${fmt(v)}`,n==="cash"?"เงินสด":"QR"]}/>
              <Legend formatter={v=>v==="cash"?"เงินสด":"QR"}/>
              <Bar dataKey="cash" stackId="a" fill="#0d9488"/>
              <Bar dataKey="qr"   stackId="a" fill="#3b82f6" radius={[5,5,0,0]}/>
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="card">
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
            <div style={{fontSize:13,color:"#5b7186",fontWeight:600}}>รายได้ต่อเครื่อง แบ่งตามเดือน</div>
            <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
              {months.map((m,i)=>(
                <span key={m} style={{fontSize:11,color:"#64748b",display:"flex",alignItems:"center",gap:4}}>
                  <span style={{width:10,height:10,borderRadius:3,background:MCOLS[i%MCOLS.length],display:"inline-block"}}/>{ml(m)}
                </span>
              ))}
            </div>
          </div>
          <ResponsiveContainer width="100%" height={Math.max(280,devices.length*17)}>
            <BarChart data={stackData} layout="vertical" barSize={9} margin={{left:0,right:14}}>
              <CartesianGrid strokeDasharray="3 3" stroke="#ffffff" horizontal={false}/>
              <XAxis type="number" tick={{fill:"#64748b",fontSize:10}} tickFormatter={v=>`฿${v}`}/>
              <YAxis type="category" dataKey="device" width={110} tick={{fill:"#64748b",fontSize:10}}/>
              <Tooltip contentStyle={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:9,color:"#1e2a3a"}} formatter={(v,n)=>v>0?[`฿${fmt(v)}`,n]:null} labelFormatter={l=>`เครื่อง ${l}`}/>
              {months.map((m,i)=>(
                <Bar key={m} dataKey={ml(m)} stackId="a" fill={MCOLS[i%MCOLS.length]} radius={i===months.length-1?[0,4,4,0]:[0,0,0,0]}/>
              ))}
            </BarChart>
          </ResponsiveContainer>
        </div>
      </>}

      {/* ── DETAIL ── */}
      {view==="detail"&&<>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14,flexWrap:"wrap"}}>
          <label style={{fontSize:13,color:"#64748b"}}>เลือกเครื่อง:</label>
          <select value={selDev} onChange={e=>setSelDev(e.target.value)}>
            {devices.map(d=><option key={d} value={d}>{d}</option>)}
          </select>
          <div style={{background:"rgba(45,212,191,0.08)",border:"1px solid rgba(45,212,191,0.25)",borderRadius:8,padding:"5px 12px",fontSize:13,color:"#0d9488"}}>
            รายได้รวม ฿{fmt(devData.reduce((s,r)=>s+r.total,0))}
          </div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(145px,1fr))",gap:10,marginBottom:14}}>
          {devData.map(m=>(
            <div key={m.month} className="sc" style={{background:"rgba(45,212,191,0.08)",border:"1px solid rgba(45,212,191,0.25)"}}>
              <div style={{fontSize:11,color:"#0d9488",fontWeight:600,marginBottom:4}}>{m.month}</div>
              <div style={{fontSize:20,fontWeight:700,color:"#0f1824"}}>฿{fmt(m.total)}</div>
              <div style={{fontSize:11,color:"#64748b",marginTop:2}}>{fmt(m.count)} ครั้ง</div>
              <div style={{fontSize:10,color:"#64748b",marginTop:1}}>เงินสด ฿{fmt(m.cash)} / QR ฿{fmt(m.qr)}</div>
            </div>
          ))}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
          <div className="card">
            <div style={{fontSize:13,color:"#5b7186",fontWeight:600,marginBottom:10}}>รายได้รายเดือน</div>
            <ResponsiveContainer width="100%" height={190}>
              <BarChart data={devData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#ffffff"/>
                <XAxis dataKey="month" tick={{fill:"#64748b",fontSize:11}}/>
                <YAxis tick={{fill:"#64748b",fontSize:11}}/>
                <Tooltip contentStyle={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:9,color:"#1e2a3a"}} formatter={(v,n)=>[`฿${fmt(v)}`,n==="cash"?"เงินสด":"QR"]}/>
                <Bar dataKey="cash" stackId="a" fill="#0d9488" name="cash"/>
                <Bar dataKey="qr"   stackId="a" fill="#3b82f6" radius={[5,5,0,0]} name="qr"/>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="card">
            <div style={{fontSize:13,color:"#5b7186",fontWeight:600,marginBottom:10}}>สัดส่วนชำระ</div>
            <ResponsiveContainer width="100%" height={170}>
              <PieChart>
                <Pie data={pieData} cx="50%" cy="50%" innerRadius={42} outerRadius={66} dataKey="value" paddingAngle={3}>
                  {pieData.map((_,i)=><Cell key={i} fill={i===0?"#0d9488":"#3b82f6"}/>)}
                </Pie>
                <Tooltip contentStyle={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:9,color:"#1e2a3a"}} formatter={v=>`฿${fmt(v)}`}/>
                <Legend/>
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
        <div className="card">
          <div style={{fontSize:13,color:"#5b7186",fontWeight:600,marginBottom:10}}>จำนวนครั้งต่อเดือน</div>
          <ResponsiveContainer width="100%" height={150}>
            <LineChart data={devData}>
              <CartesianGrid strokeDasharray="3 3" stroke="#ffffff"/>
              <XAxis dataKey="month" tick={{fill:"#64748b",fontSize:11}}/>
              <YAxis tick={{fill:"#64748b",fontSize:11}}/>
              <Tooltip contentStyle={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:9,color:"#1e2a3a"}}/>
              <Line type="monotone" dataKey="count" stroke="#a78bfa" strokeWidth={2} dot={{r:4,fill:"#a78bfa"}} name="ครั้ง"/>
            </LineChart>
          </ResponsiveContainer>
        </div>
      </>}

      {/* ── RANKING ── */}
      {view==="ranking"&&(<>
        {/* P/L tier overview */}
        <div style={{marginBottom:14}}>
          <div style={{fontSize:12,color:"#5b7186",fontWeight:600,marginBottom:8}}>ภาพรวมกำไร/ขาดทุน (P/L = AVG − ค่าเช่า)</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(120px,1fr))",gap:10}}>
            {[
              {key:"loss", label:"ขาดทุน",        sub:"P/L < 0",       color:"#dc2626", bg:"rgba(220,38,38,0.07)",  bd:"rgba(220,38,38,0.22)"},
              {key:"t0",   label:"0 – 1,000",     sub:"บาท",           color:"#64748b", bg:"rgba(148,163,184,0.09)",bd:"rgba(148,163,184,0.25)"},
              {key:"t1",   label:"1,001 – 2,000", sub:"บาท",           color:"#0d9488", bg:"rgba(45,212,191,0.08)", bd:"rgba(45,212,191,0.25)"},
              {key:"t2",   label:"2,001 – 3,000", sub:"บาท",           color:"#0d9488", bg:"rgba(45,212,191,0.11)", bd:"rgba(45,212,191,0.3)"},
              {key:"t3",   label:"มากกว่า 3,000", sub:"บาท",           color:"#0d9488", bg:"rgba(45,212,191,0.15)", bd:"rgba(45,212,191,0.38)"},
            ].map(s=>(
              <div key={s.key} className="sc" style={{background:s.bg,border:`1px solid ${s.bd}`,textAlign:"center"}}>
                <div style={{fontSize:11,fontWeight:600,color:s.color}}>{s.label}</div>
                <div style={{fontSize:26,fontWeight:700,color:"#0f1824",lineHeight:1.2}}>{plTiers[s.key]}</div>
                <div style={{fontSize:10,color:"#64748b"}}>ตู้ · {s.sub}</div>
              </div>
            ))}
          </div>
        </div>
        <div className="card">
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10,flexWrap:"wrap",gap:8}}>
            <div>
              <div style={{fontSize:13,color:"#5b7186",fontWeight:600}}>จัดอันดับรายได้ทุกเครื่อง</div>
              <div style={{fontSize:11,color:"#64748b",marginTop:2}}>รวม ฿{fmt(grandTotal)} · {rankData.length} เครื่อง</div>
            </div>
            <div style={{display:"flex",flexWrap:"wrap",gap:8,fontSize:11,color:"#64748b",alignItems:"center"}}>
              <span><span style={{color:"#0d9488"}}>▲</span> ขึ้น</span>
              <span><span style={{color:"#ef4444"}}>▼</span> ลง</span>
              <span style={{background:"rgba(220,38,38,0.12)",border:"1px solid rgba(220,38,38,0.3)",borderRadius:6,padding:"2px 8px",color:"#f87171"}}>แดง = AVG &lt; ค่าเช่า (เฉพาะ ≥2 เดือน)</span>
              <button
                onClick={async()=>{
                  setExporting(true);
                  try { await exportRankingExcel(rankData,months,rentMap,locMap,install,elapsed,lastDays,ml); }
                  catch(e){ alert("Export ไม่สำเร็จ: "+e.message); }
                  setExporting(false);
                }}
                disabled={exporting||!rankData.length}
                style={{display:"inline-flex",alignItems:"center",gap:5,padding:"5px 12px",
                  borderRadius:8,border:"1px solid rgba(45,212,191,0.4)",
                  background:exporting?"rgba(45,212,191,0.05)":"rgba(45,212,191,0.1)",
                  color:exporting||!rankData.length?"#64748b":"#0d9488",
                  cursor:exporting||!rankData.length?"not-allowed":"pointer",
                  fontSize:11,fontWeight:600,fontFamily:"inherit",transition:"all .15s"}}>
                {exporting?"⏳ กำลัง export...":"📥 Export Excel"}
              </button>
            </div>
          </div>
          <div style={{overflowX:"auto",overflowY:"auto",maxHeight:"68vh"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr style={{borderBottom:"2px solid #d8e0e8"}}>
                  {(()=>{
                    const sortable = [
                      {key:"#",     label:"#",         col:null,     align:"left"},
                      {key:"dev",   label:"เครื่อง",   col:"device", align:"left"},
                      {key:"loc",   label:"สถานที่",    col:"loc",    align:"left"},
                      {key:"inst",  label:"วันติดตั้ง", col:null,     align:"center"},
                      {key:"rent",  label:"ค่าเช่า",    col:"rent",   align:"right"},
                      ...months.map((m,i)=>({key:m, label:ml(m)+(i===months.length-1&&elapsed<lastDays?" *":""), col:null, align:"right"})),
                      {key:"avg",   label:"AVG",        col:"avg",    align:"right"},
                      {key:"pl",    label:"P/L",        col:"pl",     align:"right"},
                      {key:"total", label:"รวม",        col:"total",  align:"right"},
                      {key:"count", label:"ครั้ง",      col:"count",  align:"right"},
                      {key:"share", label:"สัดส่วน",    col:null,     align:"right"},
                      {key:"photo", label:"🖼️",          col:null,     align:"center"},
                      {key:"map",   label:"📍",          col:null,     align:"center"},
                    ];
                    return sortable.map(({key,label,col,align})=>{
                      const active = rankSort.col===col;
                      const arrow = active?(rankSort.dir==="desc"?"▼":"▲"):"⇅";
                      return (
                        <th key={key}
                          onClick={col?()=>setRankSort(s=>s.col===col?{col,dir:s.dir==="desc"?"asc":"desc"}:{col,dir:"desc"}):undefined}
                          style={{padding:"7px 9px",textAlign:align,color:active?"#0d9488":"#64748b",fontWeight:600,
                            whiteSpace:"nowrap",cursor:col?"pointer":"default",userSelect:"none",
                            position:"sticky",top:0,zIndex:2,
                            background:active?"#e8edf2":"#ffffff",
                            boxShadow:"0 2px 0 #d8e0e8",
                            transition:"color .15s"}}>
                          {label}{col&&<span style={{marginLeft:3,fontSize:9,opacity:active?1:0.4}}>{arrow}</span>}
                        </th>
                      );
                    });
                  })()}
                </tr>
              </thead>
              <tbody>
                {rankData.map((r,i)=>{
                  const prevMature = i>0 ? rankData[i-1].mature : true;
                  const showDivider = !r.mature && prevMature && i>0;
                  const isSel=r.device===selDev,rent4row=rentMap[r.device]||0,isLow=r.mature&&rent4row>0&&r.avg<rent4row;
                  return (
                    <React.Fragment key={r.device}>
                    {showDivider && (
                      <tr key="__divider__"><td colSpan={99} style={{padding:"10px",background:"#fff7e6",fontSize:12.5,color:"#b45309",fontWeight:700,textAlign:"center",letterSpacing:"0.3px",borderTop:"2px solid #f3d19e",borderBottom:"2px solid #f3d19e"}}>
                        ── ติดตั้งน้อยกว่า 2 เดือน (ยังไม่นำมาประเมิน) ──
                      </td></tr>
                    )}
                    <tr className={`rr${isSel?" rr-sel":isLow?" rr-low":""}`}>
                      <td style={{padding:"7px 9px",color:i<3?"#0d9488":"#64748b",fontWeight:700}}>#{i+1}</td>
                      <td style={{padding:"7px 9px",fontWeight:600,color:isSel?"#0d9488":isLow?"#f87171":"#1e2a3a",whiteSpace:"nowrap"}}>
                        <div style={{display:"flex",alignItems:"center",gap:5}}>
                          {r.device}
                          <span style={{fontSize:9,background:r.mature?"rgba(100,116,139,0.12)":"rgba(251,191,36,0.15)",border:r.mature?"1px solid rgba(100,116,139,0.25)":"1px solid rgba(251,191,36,0.35)",borderRadius:4,padding:"1px 6px",color:r.mature?"#64748b":"#d97706",fontWeight:600,whiteSpace:"nowrap"}}>{r.monthsInstalled} เดือน</span>
                          <button
                            onClick={e=>{
                              e.stopPropagation();
                              try {
                                const el=document.createElement("textarea");
                                el.value=r.device; el.style.position="fixed"; el.style.opacity="0";
                                document.body.appendChild(el); el.focus(); el.select();
                                document.execCommand("copy");
                                document.body.removeChild(el);
                              } catch(_) {}
                              setCopied(r.device); setTimeout(()=>setCopied(""),1500);
                            }}
                            title="คัดลอกหมายเลขเครื่อง"
                            style={{padding:"2px 7px",fontSize:11,borderRadius:5,border:"1px solid #d8e0e8",background:copied===r.device?"rgba(45,212,191,0.15)":"rgba(15,24,36,0.03)",color:copied===r.device?"#0d9488":"#64748b",cursor:"pointer",fontFamily:"inherit",lineHeight:1.5,transition:"all .15s",flexShrink:0}}>
                            {copied===r.device?"✓":"⎘"}
                          </button>
                        </div>
                      </td>
                      <td style={{padding:"7px 9px",color:"#5b7186",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={locMap[r.device]||""}>
                        {locMap[r.device]||<span style={{color:"#d8e0e8"}}>—</span>}
                      </td>
                      <td style={{padding:"7px 9px",textAlign:"center",color:"#64748b",whiteSpace:"nowrap"}}>{install[r.device]||"—"}</td>
                      <td style={{padding:"7px 9px",textAlign:"right",color:"#0f1824",fontWeight:700,whiteSpace:"nowrap"}}>฿{fmt(rentMap[r.device]||0)}</td>
                      {months.map((m,mi)=>cv(r,m,mi>0, mi===months.length-1&&elapsed<lastDays))}
                      <td style={{padding:"7px 9px",textAlign:"right",whiteSpace:"nowrap"}}>
                        <span style={{color:isLow?"#f87171":r.avg>=(rentMap[r.device]||0)?"#0d9488":"#5b7186",fontWeight:isLow?700:400}}>฿{fmt(r.avg)}</span>
                      </td>
                      <td style={{padding:"7px 9px",textAlign:"right",whiteSpace:"nowrap"}}>
                        {(()=>{const pl=r.avg-(rentMap[r.device]||0);return <span style={{color:pl>=0?"#0d9488":"#ef4444",fontWeight:700}}>{pl>=0?"+":"−"}฿{fmt(Math.abs(pl))}</span>;})()}
                      </td>
                      <td style={{padding:"7px 9px",textAlign:"right",fontWeight:700,color:"#0d9488",whiteSpace:"nowrap"}}>฿{fmt(r.total)}</td>
                      <td style={{padding:"7px 9px",textAlign:"right",color:"#64748b"}}>{fmt(r.count)}</td>
                      <td style={{padding:"7px 9px",textAlign:"right"}}>
                        <div style={{display:"flex",alignItems:"center",justifyContent:"flex-end",gap:5}}>
                          <div style={{width:46,height:4,background:"#e6ebf1",borderRadius:3,overflow:"hidden"}}>
                            <div style={{width:`${Math.min(100,r.total/rankData[0].total*100)}%`,height:"100%",background:isLow?"#ef4444":"#0d9488",borderRadius:3}}/>
                          </div>
                          <span style={{fontSize:10,color:"#64748b",minWidth:30}}>{(r.total/grandTotal*100).toFixed(1)}%</span>
                        </div>
                      </td>
                      <td style={{padding:"7px 9px",textAlign:"center"}}>
                        {photoMap[r.device]
                          ? <PhotoCell url={photoMap[r.device]} label={locMap[r.device]||r.device} />
                          : <span style={{color:"#d8e0e8",fontSize:11}}>—</span>}
                      </td>
                      <td style={{padding:"7px 9px",textAlign:"center"}}>
                        {mapUrl[r.device]
                          ? <button onClick={e=>{e.stopPropagation(); setMapModal({label:locMap[r.device]||r.device});}}
                              style={{display:"inline-flex",alignItems:"center",gap:3,background:"rgba(45,212,191,0.08)",border:"1px solid rgba(45,212,191,0.25)",borderRadius:6,padding:"3px 8px",color:"#0d9488",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>
                              📍
                            </button>
                          : <span style={{color:"#d8e0e8",fontSize:11}}>—</span>}
                      </td>
                    </tr>
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
          <p style={{fontSize:10,color:"#d8e0e8",marginTop:8,lineHeight:1.5}}>* = ค่า RunRate (ประมาณการ {elapsed}/{lastDays} วัน) · ตัวเลขสีน้ำเงิน = RunRate · AVG เฉลี่ยทุกเดือน (เดือน * ใช้ RunRate) · แดง = AVG &lt; ค่าเช่า · เรียงตาม P/L มาก→น้อย</p>
        </div>
      </>)}

      {/* ── DECISION ── */}
      {view==="decision"&&(
        <div>
          {/* Executive summary box — copy & send on LINE */}
          <div className="card" style={{marginBottom:14,borderColor:"rgba(45,212,191,0.35)"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10,marginBottom:10}}>
              <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                <span style={{fontSize:13,fontWeight:700,color:"#0d9488"}}>📋 สรุปตู้สำหรับส่งผู้บริหาร (LINE)</span>
                <select value={sumMonth} onChange={e=>setSumMonth(e.target.value)}>
                  {months.map(m=>(
                    <option key={m} value={m}>{ml(m)}{m===lastM&&elapsed<lastDays?" (ยังไม่จบเดือน)":""}</option>
                  ))}
                </select>
              </div>
              <button onClick={copySummary}
                style={{display:"inline-flex",alignItems:"center",gap:6,padding:"7px 16px",borderRadius:8,
                  border:"1px solid rgba(45,212,191,0.4)",
                  background:copiedSum?"rgba(45,212,191,0.2)":"rgba(45,212,191,0.1)",
                  color:"#0d9488",fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit",transition:"all .15s"}}>
                {copiedSum?"✓ คัดลอกแล้ว":"📋 คัดลอกข้อความ"}
              </button>
            </div>
            <textarea readOnly value={summaryText} onFocus={e=>e.target.select()}
              style={{width:"100%",minHeight:170,resize:"vertical",boxSizing:"border-box",
                background:"#f8fafc",border:"1px solid #d8e0e8",borderRadius:10,padding:"12px 14px",
                fontFamily:"'IBM Plex Sans Thai','Sarabun',monospace",fontSize:14,lineHeight:1.7,
                color:"#1e2a3a",outline:"none",whiteSpace:"pre"}}/>
            <div style={{fontSize:10.5,color:"#64748b",marginTop:6}}>แบ่งกลุ่มตามยอดสุทธิ (ยอดขาย − ค่าเช่า) · นับเฉพาะตู้ที่มียอดในเดือนนั้น · ตู้ที่ยังไม่มีข้อมูลค่าเช่าจะคิดค่าเช่า = 0 · แตะช่องเพื่อเลือกทั้งหมด แล้วคัดลอกไปวางในไลน์ได้เลย</div>
          </div>

          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(145px,1fr))",gap:10,marginBottom:16}}>
            {[
              {key:"keep",    icon:"✅",label:"วางตู้ต่อ",      color:"#0d9488",bg:"rgba(45,212,191,0.07)",  bd:"rgba(45,212,191,0.2)"},
              {key:"watch",   icon:"👀",label:"เฝ้าดู",          color:"#d97706",bg:"rgba(251,191,36,0.07)",   bd:"rgba(251,191,36,0.2)"},
              {key:"relocate",icon:"🔴",label:"ย้ายออก",      color:"#dc2626",bg:"rgba(220,38,38,0.08)",    bd:"rgba(220,38,38,0.2)"},
              {key:"nodata",  icon:"⚠️",label:"ตรวจสอบสถานะ",  color:"#64748b",bg:"rgba(148,163,184,0.07)",bd:"rgba(148,163,184,0.2)"},
            ].map(s=>(
              <div key={s.key} className="sc" style={{background:s.bg,border:`1px solid ${s.bd}`}}>
                <div style={{fontSize:20,marginBottom:5}}>{s.icon}</div>
                <div style={{fontSize:28,fontWeight:700,color:s.color}}>{grp[s.key].length}</div>
                <div style={{fontSize:13,color:"#64748b",marginTop:2}}>{s.label}</div>
                <div style={{fontSize:10,color:"#64748b"}}>เครื่อง</div>
              </div>
            ))}
          </div>

          <div className="card" style={{marginBottom:12,padding:"11px 14px"}}>
            <div style={{fontSize:11,color:"#64748b",fontWeight:600,marginBottom:7}}>เกณฑ์การตัดสิน</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(195px,1fr))",gap:5,fontSize:11,color:"#64748b"}}>
              <span>✅ <b style={{color:"#0d9488"}}>วางตู้ต่อ</b> — Run Rate ≥ ค่าเช่า และ AVG ≥ 80%</span>
              <span>👀 <b style={{color:"#f59e0b"}}>เฝ้าดู</b> — ผ่านเกณฑ์อย่างใดอย่างหนึ่ง</span>
              <span>🔴 <b style={{color:"#ef4444"}}>ย้ายออก</b> — ทั้ง Run Rate และ AVG ต่ำกว่าค่าเช่า</span>
              <span>⚠️ <b style={{color:"#64748b"}}>ตรวจสอบ</b> — ไม่มีข้อมูลเดือนล่าสุด</span>
            </div>
          </div>

          {[
            {key:"relocate",label:"🔴 ย้ายออก",        color:"#f87171",bd:"rgba(220,38,38,0.3)"},
            {key:"watch",   label:"👀 เฝ้าดู / พิจารณา",  color:"#fbbf24",bd:"rgba(251,191,36,0.3)"},
            {key:"keep",    label:"✅ วางตู้ต่อ",          color:"#0d9488",bd:"#0d9488"},
            {key:"nodata",  label:"⚠️ ตรวจสอบสถานะ",     color:"#64748b",bd:"rgba(148,163,184,0.15)"},
          ].map(({key,label,color,bd})=>{
            const g=grp[key]; if(!g.length) return null;
            return (
              <div key={key} className="card" style={{marginBottom:12,borderColor:bd}}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:9}}>
                  <span style={{fontWeight:700,color,fontSize:14}}>{label}</span>
                  <span style={{fontSize:11,color:"#64748b"}}>{g.length} เครื่อง</span>
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead>
                      <tr style={{borderBottom:"1px solid #e6ebf1"}}>
                        {[["เครื่อง","left"],["สถานที่","left"],["วันติดตั้ง","center"],["ค่าเช่า","right"],["AVG","right"],["Run Rate","right"],["เหตุผล","left"]].map(([h,a])=>(
                          <th key={h} style={{padding:"6px 9px",textAlign:a,color:"#64748b",fontWeight:600}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {g.map(r=>(
                        <tr key={r.device} className="rr" onClick={()=>{setSelDev(r.device);setView("detail");}}
                          style={{borderBottom:"1px solid #f1f5f9"}}>
                          <td style={{padding:"7px 9px",fontWeight:600,color:"#1e2a3a",whiteSpace:"nowrap"}}>{r.device}</td>
                          <td style={{padding:"7px 9px",color:"#5b7186",maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={locMap[r.device]||""}>{locMap[r.device]||<span style={{color:"#d8e0e8"}}>—</span>}</td>
                          <td style={{padding:"7px 9px",textAlign:"center",color:"#64748b",whiteSpace:"nowrap"}}>{install[r.device]||"—"}</td>
                          <td style={{padding:"7px 9px",textAlign:"right",color:"#0f1824",fontWeight:700,whiteSpace:"nowrap"}}>฿{fmt(r.rent)}</td>
                          <td style={{padding:"7px 9px",textAlign:"right",whiteSpace:"nowrap"}}>
                            <span style={{color:r.avg>=r.rent?"#0d9488":"#ef4444",fontWeight:600}}>฿{fmt(r.avg)}</span>
                          </td>
                          <td style={{padding:"7px 9px",textAlign:"right",whiteSpace:"nowrap"}}>
                            {r.rr>0?<span style={{color:r.rr>=r.rent?"#0d9488":"#ef4444",fontWeight:600}}>฿{fmt(r.rr)}</span>:<span style={{color:"#d8e0e8"}}>—</span>}
                          </td>
                          <td style={{padding:"7px 9px",color:"#64748b",fontSize:11}}>{r.reason}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })}
        </div>
      )}
      <MapModal item={mapModal} onClose={()=>setMapModal(null)}/>
    </div>
  );
}
