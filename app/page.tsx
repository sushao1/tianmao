"use client";

import { useMemo, useRef, useState } from "react";
import JSZip from "jszip";
import * as XLSX from "xlsx";

type Review = {
  id: string;
  sourceRow: number;
  nickname: string;
  sku: string;
  rating: number;
  tags: string[];
  text: string;
  followText: string;
  date: string;
  likes: number;
  replies: number;
  images: string[];
  followImages: string[];
};

type ParsedImage = {
  name: string;
  url: string;
};

const sampleReviews: Review[] = [
  {
    id: "sample-1",
    sourceRow: 1,
    nickname: "爱**购",
    sku: "颜色分类：白茶清欢",
    rating: 5,
    tags: ["物有所值"],
    text:
      "实物很高级，最重要的香味我很喜欢，能选到一款好闻的香薰还是挺不容易的，这个价位很值",
    followText: "用了几天香味还是挺稳的，不刺鼻，放车里刚刚好。",
    date: "2026-07-28",
    likes: 10,
    replies: 0,
    images: [],
    followImages: [],
  },
  {
    id: "sample-2",
    sourceRow: 2,
    nickname: "李**33",
    sku: "颜色分类：空谷幽兰",
    rating: 5,
    tags: ["颜值高"],
    text: "这款的设计我超喜欢，瓶身有新中式的感觉，味道也好闻，很绝～",
    followText: "",
    date: "2026-07-26",
    likes: 7,
    replies: 0,
    images: [],
    followImages: [],
  },
  {
    id: "sample-3",
    sourceRow: 3,
    nickname: "果**然",
    sku: "颜色分类：白茶清欢",
    rating: 5,
    tags: ["香味不刺鼻"],
    text:
      "买来放车里用的，女儿有鼻炎，之前一直没选到合适的，这一款终于过关了，淡淡的不会冲。",
    followText: "后面又放了一周，车里味道还是淡淡的，没有闷人的感觉。",
    date: "2026-07-24",
    likes: 4,
    replies: 1,
    images: [],
    followImages: [],
  },
];

const fieldAliases = {
  nickname: ["昵称", "买家", "用户", "用户昵称", "name", "nickname"],
  sku: ["规格", "sku", "SKU", "颜色分类", "型号", "属性"],
  rating: ["评分", "星级", "rating", "rate"],
  tags: ["标签", "评价标签", "印象", "关键词", "tag", "tags"],
  text: ["评价", "评价内容", "内容", "初评", "文字", "评论", "review", "text"],
  date: ["日期", "评价时间", "时间", "date", "created_at"],
  likes: ["点赞", "有用", "赞", "likes"],
  replies: ["回复", "评论数", "reply", "replies"],
  image: ["图片", "图片名", "图片文件", "image", "images", "pic", "photo"],
};

const textAliases = [...fieldAliases.text, "追评", "后续追评"];
const followTextAliases = ["追评", "后续追评", "追加评价", "追加评论", "追评内容", "after", "follow"];

function findValue(row: Record<string, unknown>, aliases: string[]) {
  const keys = Object.keys(row);
  for (const alias of aliases) {
    const key = keys.find((item) => item.trim().toLowerCase() === alias.toLowerCase());
    if (key && row[key] !== undefined && row[key] !== null) return String(row[key]).trim();
  }
  const looseKey = keys.find((key) =>
    aliases.some((alias) => key.toLowerCase().includes(alias.toLowerCase())),
  );
  if (looseKey && row[looseKey] !== undefined && row[looseKey] !== null) {
    return String(row[looseKey]).trim();
  }
  return "";
}

function fallbackNickname(index: number) {
  const names = ["爱**购", "李**33", "果**然", "小**呀", "清**风", "买**啦", "淘**友"];
  return names[index % names.length];
}

function parseTags(value: string) {
  return value
    .split(/[、,，;；\n]/)
    .map((item) => item.trim())
    .filter(Boolean)
    .slice(0, 4);
}

function parseRating(value: string) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return 5;
  return Math.max(1, Math.min(5, Math.round(parsed)));
}

function fileBase(name: string) {
  return name.replace(/\.[^.]+$/, "").trim().toLowerCase();
}

function cellRow(cellRef: string) {
  const match = cellRef.match(/\d+/);
  return match ? Number(match[0]) : 0;
}

function cellColumn(cellRef: string) {
  const letters = cellRef.match(/[A-Z]+/i)?.[0].toUpperCase() || "";
  return Array.from(letters).reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
}

function mimeFromPath(path: string) {
  const ext = path.split(".").pop()?.toLowerCase();
  if (ext === "jpg" || ext === "jpeg") return "image/jpeg";
  if (ext === "png") return "image/png";
  if (ext === "webp") return "image/webp";
  if (ext === "gif") return "image/gif";
  return "application/octet-stream";
}

function xmlDoc(text: string) {
  return new DOMParser().parseFromString(text, "application/xml");
}

function nodesByLocalName(root: Document | Element, name: string) {
  return Array.from(root.getElementsByTagName("*")).filter((node) => node.localName === name);
}

function attrByLocalName(node: Element, name: string) {
  for (const attr of Array.from(node.attributes)) {
    if (attr.localName === name) return attr.value;
  }
  return "";
}

async function extractWorkbookImages(buffer: ArrayBuffer) {
  const zip = await JSZip.loadAsync(buffer);
  const allImages: ParsedImage[] = [];
  const rowImages = new Map<number, string[]>();
  const rowColumnImages = new Map<number, Array<{ column: number; url: string }>>();
  const idToMedia = new Map<string, string>();
  const relToTarget = new Map<string, string>();

  await Promise.all(
    Object.keys(zip.files)
      .filter((path) => /^xl\/media\/.+\.(png|jpe?g|webp|gif)$/i.test(path))
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
      .map(async (path) => {
        const file = zip.file(path);
        if (!file) return;
        const blob = await file.async("blob");
        allImages.push({ name: path.split("/").pop() || path, url: URL.createObjectURL(blob) });
      }),
  );

  const relsFile = zip.file("xl/_rels/cellimages.xml.rels");
  if (relsFile) {
    const rels = xmlDoc(await relsFile.async("text"));
    nodesByLocalName(rels, "Relationship").forEach((node) => {
      const element = node as Element;
      const id = element.getAttribute("Id") || "";
      const target = element.getAttribute("Target") || "";
      if (id && target) relToTarget.set(id, target.startsWith("xl/") ? target : `xl/${target}`);
    });
  }

  const cellImagesFile = zip.file("xl/cellimages.xml");
  if (cellImagesFile) {
    const cellImages = xmlDoc(await cellImagesFile.async("text"));
    nodesByLocalName(cellImages, "pic").forEach((pic) => {
      const nameNode = nodesByLocalName(pic, "cNvPr")[0] as Element | undefined;
      const blipNode = nodesByLocalName(pic, "blip")[0] as Element | undefined;
      const imageId = nameNode?.getAttribute("name") || "";
      const relId = blipNode ? attrByLocalName(blipNode, "embed") : "";
      const mediaPath = relToTarget.get(relId);
      if (imageId && mediaPath) idToMedia.set(imageId, mediaPath);
    });
  }

  const mediaUrlCache = new Map<string, string>();
  async function mediaUrl(path: string) {
    const cached = mediaUrlCache.get(path);
    if (cached) return cached;
    const file = zip.file(path);
    if (!file) return "";
    const blob = await file.async("blob");
    const url = URL.createObjectURL(new Blob([blob], { type: mimeFromPath(path) }));
    mediaUrlCache.set(path, url);
    return url;
  }

  const sheetNames = Object.keys(zip.files).filter((path) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(path));
  for (const sheetPath of sheetNames) {
    const sheet = xmlDoc(await zip.file(sheetPath)!.async("text"));
    for (const cell of nodesByLocalName(sheet, "c")) {
      const element = cell as Element;
      const ref = element.getAttribute("r") || "";
      const rowNumber = cellRow(ref);
      const columnNumber = cellColumn(ref);
      if (!rowNumber) continue;
      const formula = nodesByLocalName(element, "f")[0]?.textContent || "";
      const ids = Array.from(formula.matchAll(/DISPIMG\(["&quot;]*([^",)&]+)["&quot;]*/g)).map((match) => match[1]);
      for (const id of ids) {
        const path = idToMedia.get(id);
        if (!path) continue;
        const url = await mediaUrl(path);
        if (!url) continue;
        const existing = rowImages.get(rowNumber) || [];
        if (!existing.includes(url)) rowImages.set(rowNumber, [...existing, url]);
        const existingColumnImages = rowColumnImages.get(rowNumber) || [];
        if (!existingColumnImages.some((item) => item.url === url)) {
          rowColumnImages.set(rowNumber, [...existingColumnImages, { column: columnNumber, url }]);
        }
      }
    }
  }

  return { allImages, rowImages, rowColumnImages };
}

function matchImages(imageValue: string, index: number, images: ParsedImage[]) {
  const refs = imageValue
    .split(/[、,，;；\s]+/)
    .map((item) => fileBase(item))
    .filter(Boolean);

  if (refs.length) {
    const matched = images
      .filter((image) => refs.some((ref) => fileBase(image.name).includes(ref) || ref.includes(fileBase(image.name))))
      .map((image) => image.url);
    if (matched.length) return matched.slice(0, 9);
  }

  return images[index] ? [images[index].url] : [];
}

function sortNewestFirst(reviews: Review[]) {
  return [...reviews].sort((a, b) => b.sourceRow - a.sourceRow);
}

function imageColumnGroups(headers: string[]) {
  const normalized = headers.map((header) => header.trim());
  const followIndex = normalized.findIndex((header) => followTextAliases.includes(header));
  const commentIndex = normalized.findIndex((header, index) => followIndex >= 0 && index > followIndex && header.includes("评中评"));
  const initialImageHeader = normalized.findIndex((header) => header === "图片");
  const followImageHeader = normalized.findIndex((header, index) => followIndex >= 0 && index > followIndex && header === "图片");

  return {
    initialStart: initialImageHeader >= 0 ? initialImageHeader + 1 : 0,
    initialEnd: followIndex >= 0 ? followIndex + 1 : headers.length + 1,
    followStart: followImageHeader >= 0 ? followImageHeader + 1 : 0,
    followEnd: commentIndex >= 0 ? commentIndex + 1 : headers.length + 1,
  };
}

function splitRowImages(
  rowItems: Array<{ column: number; url: string }> | undefined,
  groups: ReturnType<typeof imageColumnGroups>,
) {
  const initialImages: string[] = [];
  const followImages: string[] = [];

  for (const item of rowItems || []) {
    if (groups.followStart && item.column >= groups.followStart && item.column < groups.followEnd) {
      followImages.push(item.url);
    } else if (groups.initialStart && item.column >= groups.initialStart && item.column < groups.initialEnd) {
      initialImages.push(item.url);
    }
  }

  return { initialImages, followImages };
}

function makeReview(row: Record<string, unknown>, index: number, images: ParsedImage[], sourceRow: number): Review {
  const text = findValue(row, fieldAliases.text);
  const followText = findValue(row, followTextAliases);
  const date = findValue(row, fieldAliases.date);
  const tagValue = findValue(row, fieldAliases.tags);
  const imageValue = findValue(row, fieldAliases.image);

  return {
    id: `review-${index}-${Math.random().toString(16).slice(2)}`,
    sourceRow,
    nickname: findValue(row, fieldAliases.nickname) || fallbackNickname(index),
    sku: findValue(row, fieldAliases.sku) || "颜色分类：默认规格",
    rating: parseRating(findValue(row, fieldAliases.rating)),
    tags: parseTags(tagValue || (index % 2 ? "颜值高" : "物有所值")),
    text: text || "这条评价没有识别到文字内容，请检查表格里的“评价内容/初评/文字”列。",
    followText,
    date: date || "2026-08-26",
    likes: Number(findValue(row, fieldAliases.likes)) || Math.max(1, 10 - index),
    replies: Number(findValue(row, fieldAliases.replies)) || 0,
    images: matchImages(imageValue, index, images),
    followImages: [],
  };
}

function rowToObject(headers: string[], row: unknown[]) {
  const object: Record<string, unknown> = {};
  headers.forEach((header, index) => {
    const key = header || `列${index + 1}`;
    if (object[key] === undefined) {
      object[key] = row[index] ?? "";
    } else {
      object[`${key}_${index + 1}`] = row[index] ?? "";
    }
  });
  return object;
}

function detectHeaderIndex(rows: unknown[][]) {
  return rows.findIndex((row) =>
    row.some((value) => textAliases.some((alias) => String(value).trim().toLowerCase() === alias.toLowerCase())),
  );
}

function displayWidth(char: string) {
  return 1;
}

function wrapTaobaoLine(text: string, maxWidth = 20) {
  return text
    .split(/\r?\n/)
    .map((line) => {
      const chars = Array.from(line.trim());
      const chunks: string[] = [];
      let current = "";
      let currentWidth = 0;

      for (const char of chars) {
        const width = displayWidth(char);
        if (current && currentWidth + width > maxWidth) {
          chunks.push(current);
          current = char;
          currentWidth = width;
        } else {
          current += char;
          currentWidth += width;
        }
      }

      if (current) chunks.push(current);
      return chunks.join("\n");
    })
    .join("\n");
}

function StarRating({ rating }: { rating: number }) {
  return <div className="stars" aria-label={`${rating}星`}>{"★★★★★".slice(0, rating)}</div>;
}

function Avatar({ nickname, index }: { nickname: string; index: number }) {
  const colors = ["#75d9d5", "#9b7b59", "#79d7bd", "#ffb86b", "#8aa8ff"];
  return (
    <div className="avatar" style={{ background: colors[index % colors.length] }}>
      {nickname.slice(0, 1)}
    </div>
  );
}

function ReviewCard({
  review,
  index,
  mode = "normal",
}: {
  review: Review;
  index: number;
  mode?: "normal" | "follow";
}) {
  const imageBlock = !!review.images.length && (
    <div className={`image-grid count-${Math.min(review.images.length, 3)}`}>
      {review.images.slice(0, 9).map((src, imageIndex) => (
        <img src={src} alt={`${review.nickname} 评价图 ${imageIndex + 1}`} key={src + imageIndex} />
      ))}
    </div>
  );
  const followImageBlock = !!review.followImages.length && (
    <div className={`image-grid count-${Math.min(review.followImages.length, 3)}`}>
      {review.followImages.slice(0, 9).map((src, imageIndex) => (
        <img src={src} alt={`${review.nickname} 追评图 ${imageIndex + 1}`} key={src + imageIndex} />
      ))}
    </div>
  );

  return (
    <article className="review-card">
      <div className="review-head">
        <Avatar nickname={review.nickname} index={index} />
        <div>
          <div className="nickname">{review.nickname}</div>
          <div className="sku">{review.sku}</div>
        </div>
      </div>

      <StarRating rating={review.rating} />

      {!!review.tags.length && (
        <div className="review-tags">
          {review.tags.map((tag) => (
            <span key={tag}>{tag}</span>
          ))}
        </div>
      )}

      {mode === "follow" ? (
        <>
          <div className="initial-review">
            <div className="section-title">初次评价</div>
            <p className="review-text">{wrapTaobaoLine(review.text)}</p>
          </div>
          {imageBlock}
          <div className="follow-review follow-page">
            <div className="follow-title">追加评价</div>
            <p>{wrapTaobaoLine(review.followText)}</p>
            {followImageBlock}
          </div>
        </>
      ) : (
        <>
          <p className="review-text">{wrapTaobaoLine(review.text)}</p>
          {imageBlock}
        </>
      )}

      {mode !== "follow" && review.followText && (
        <div className="follow-review">
          <div className="follow-title">追加评价</div>
          <p>{wrapTaobaoLine(review.followText)}</p>
          {followImageBlock}
        </div>
      )}

      <div className="review-foot">
        <span>{review.date}</span>
        <span>👍 {review.likes}</span>
        <span>💬 回复{review.replies ? ` ${review.replies}` : ""}</span>
      </div>
    </article>
  );
}

export default function Home() {
  const [reviews, setReviews] = useState<Review[]>(sampleReviews);
  const [images, setImages] = useState<ParsedImage[]>([]);
  const [activeTab, setActiveTab] = useState("有图");
  const [notice, setNotice] = useState("可上传 Excel / CSV；Excel 内嵌图片会自动读取，追评页会同时展示初评和追加评价。");
  const inputRef = useRef<HTMLInputElement>(null);

  const stats = useMemo(() => {
    const withImage = reviews.filter((review) => review.images.length || review.followImages.length).length;
    const good = reviews.filter((review) => review.rating >= 4).length;
    const follow = reviews.filter((review) => review.followText.trim()).length;
    return { total: reviews.length, withImage, good, follow };
  }, [reviews]);

  const visibleReviews = useMemo(() => {
    if (activeTab === "有图") return reviews.filter((review) => review.images.length || review.followImages.length);
    if (activeTab === "好评") return reviews.filter((review) => review.rating >= 4);
    if (activeTab === "追评") return reviews.filter((review) => review.followText.trim());
    return reviews;
  }, [activeTab, reviews]);

  const topTags = useMemo(() => {
    const counts = new Map<string, number>();
    reviews.forEach((review) => review.tags.forEach((tag) => counts.set(tag, (counts.get(tag) || 0) + 1)));
    return Array.from(counts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8);
  }, [reviews]);

  async function handleUpload(fileList: FileList | null) {
    if (!fileList?.length) return;

    const files = Array.from(fileList);
    const imageFiles = files.filter((file) => file.type.startsWith("image/"));
    const dataFile = files.find((file) => /\.(xlsx|xls|csv)$/i.test(file.name));
    const manualImages = imageFiles.map((file) => ({ name: file.name, url: URL.createObjectURL(file) }));

    if (!dataFile) {
      setImages(manualImages);
      setReviews((current) =>
        current.map((review, index) => ({
          ...review,
          images: manualImages[index] ? [manualImages[index].url] : review.images,
        })),
      );
      setNotice(`已读取 ${manualImages.length} 张图片。继续上传 Excel/CSV 后可自动匹配到评价。`);
      return;
    }

    const buffer = await dataFile.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const sheetRows = XLSX.utils.sheet_to_json<unknown[]>(firstSheet, { header: 1, defval: "" });
    const headerIndex = detectHeaderIndex(sheetRows);
    const headers = (sheetRows[headerIndex >= 0 ? headerIndex : 0] || []).map((value) => String(value).trim());
    const embedded = /\.(xlsx|xlsm)$/i.test(dataFile.name)
      ? await extractWorkbookImages(buffer)
      : {
          allImages: [] as ParsedImage[],
          rowImages: new Map<number, string[]>(),
          rowColumnImages: new Map<number, Array<{ column: number; url: string }>>(),
        };
    const availableImages = manualImages.length ? manualImages : embedded.allImages.length ? embedded.allImages : images;
    const columnGroups = imageColumnGroups(headers);
    const dataRows = sheetRows.slice((headerIndex >= 0 ? headerIndex : 0) + 1);
    const parsedReviews = dataRows
      .map((row, index) => {
        const excelRowNumber = (headerIndex >= 0 ? headerIndex : 0) + 2 + index;
        const rowObject = rowToObject(headers, row);
        const review = makeReview(rowObject, index, availableImages, excelRowNumber);
        const splitImages = splitRowImages(embedded.rowColumnImages.get(excelRowNumber), columnGroups);
        const embeddedImages = embedded.rowImages.get(excelRowNumber);

        if (splitImages.initialImages.length || splitImages.followImages.length) {
          return {
            ...review,
            images: splitImages.initialImages.slice(0, 9),
            followImages: splitImages.followImages.slice(0, 9),
          };
        }

        return embeddedImages?.length ? { ...review, images: embeddedImages.slice(0, 9) } : review;
      })
      .filter((review) => review.text && !review.text.includes("没有识别到文字内容"));

    if (parsedReviews.length) {
      setReviews(sortNewestFirst(parsedReviews));
      if (manualImages.length) setImages(manualImages);
      if (!manualImages.length && embedded.allImages.length) setImages(embedded.allImages);
      setNotice(
        `已解析 ${parsedReviews.length} 条评价，已按“表格下方最新评价在最上面”排序；识别到 Excel 内嵌图片 ${embedded.allImages.length} 张${
          manualImages.length ? `，额外上传图片 ${manualImages.length} 张` : ""
        }。`,
      );
    } else {
      setNotice("没有从文件里识别到评价行，请确认第一行是表头，且包含评价内容/初评/文字列。");
    }
  }

  return (
    <main className="page-shell">
      <section className="control-panel">
        <div className="eyebrow">淘宝手机端评价预览</div>
        <h1>上传评价表，预览淘宝评论区效果</h1>
        <p className="lead">
          按淘宝“宝贝评价”评论区窗口预览评价文字、标签、规格、图片和追评。汉字、标点、英文和数字都按 1 个字符计算，每行最多 20 个字符，超出自动换行。
        </p>

        <div
          className="upload-box"
          role="button"
          tabIndex={0}
          onClick={() => inputRef.current?.click()}
          onKeyDown={(event) => event.key === "Enter" && inputRef.current?.click()}
        >
          <input
            ref={inputRef}
            type="file"
            multiple
            accept=".xlsx,.xls,.csv,image/*"
            onChange={(event) => handleUpload(event.target.files)}
          />
          <div className="upload-icon">⇧</div>
          <strong>上传文件更新预览</strong>
          <span>Excel / CSV / JPG / PNG，可多选</span>
        </div>

        <p className="notice">{notice}</p>

        <div className="format-card">
          <h2>推荐表头</h2>
          <div className="format-grid">
            <span>昵称</span>
            <span>规格</span>
            <span>评分</span>
            <span>标签</span>
            <span>评价内容</span>
            <span>追评</span>
            <span>日期</span>
            <span>点赞</span>
            <span>图片</span>
          </div>
          <p>Excel 里有内嵌图片时会自动读取；追评列有内容时，会在追评 tab 里和初评一起展示。</p>
        </div>

        <div className="checks">
          <h2>预览时重点看</h2>
          <ul>
            <li>前 3 条是否覆盖核心需求点。</li>
            <li>汉字、标点、英文和数字都按 1 个字符计算，每行最多 20 个字符。</li>
            <li>图片是否连续堆叠成图片柱。</li>
            <li>追评是否能和初评连起来看。</li>
          </ul>
        </div>
      </section>

      <section className="phone-stage" aria-label="淘宝手机端评价预览">
        <div className="phone">
          <div className="phone-status" />
          <header className="taobao-header">
            <span>‹</span>
            <strong>宝贝评价</strong>
            <span>…</span>
          </header>

          <div className="score-block">
            <div className="score-main">
              <strong>{stats.good ? "4.9" : "4.6"}</strong>
              <span>综合评分</span>
            </div>
            <div className="tag-cloud">
              {topTags.length ? (
                topTags.map(([tag, count]) => (
                  <span key={tag}>
                    {tag}
                    {count}
                  </span>
                ))
              ) : (
                <>
                  <span>物流快1</span>
                  <span>颜值高1</span>
                </>
              )}
            </div>
          </div>

          <nav className="tabs">
            {[
              ["全部", stats.total],
              ["有图", stats.withImage],
              ["好评", stats.good],
              ["追评", stats.follow],
            ].map(([label, count]) => (
              <button
                className={activeTab === label ? "active" : ""}
                key={label}
                onClick={() => setActiveTab(String(label))}
              >
                {label}
                <span>({count})</span>
              </button>
            ))}
          </nav>

          <div className="review-list">
            {visibleReviews.length ? (
              visibleReviews.map((review, index) => (
                <ReviewCard
                  review={review}
                  index={index}
                  key={review.id}
                  mode={activeTab === "追评" ? "follow" : "normal"}
                />
              ))
            ) : (
              <div className="empty-state">当前筛选下没有评价。请检查 Excel 对应列是否有文字或图片。</div>
            )}
          </div>
        </div>
      </section>
    </main>
  );
}
