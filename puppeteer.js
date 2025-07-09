// node puppeteer.js
// /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \         
// --remote-debugging-port=9222 \
// --user-data-dir=~/chrome-data \

// 날짜 10개씩 실행   (타임아웃)

const puppeteer = require("puppeteer-core");
const { createObjectCsvWriter } = require("csv-writer");
const fs = require("fs").promises;
const xlsx = require("xlsx");

// CSV 설정
const csvWriter = createObjectCsvWriter({
  path: "법원_사건_상세_최종.csv",
  header: [
    { id: "법원", title: "법원" },
    { id: "사건번호", title: "사건번호" },
    { id: "사건명", title: "사건명" },
    { id: "문서명", title: "문서명" },
    { id: "접수일자", title: "접수일자" },
    { id: "소송등인지", title: "소송등인지(원)" },
    { id: "송달료", title: "송달료(원)" },
    { id: "법원보관금", title: "법원보관금(원)" },
    { id: "납부금액", title: "총납부액(원)" },
    { id: "가상계좌번호", title: "가상계좌번호" },
    { id: "납부증여부", title: "납부증" },
  ],
  encoding: "utf8",
  fieldDelimiter: ",",
  append: false,
});

// 유틸리티 함수
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// 클릭 함수 개선
const safeClick = async (page, selector, options = {}) => {
  const { maxRetries = 3, retryDelay = 1000, waitAfter = 300 } = options;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const element = await page.waitForSelector(selector, {
        visible: true,
        timeout: 5000,
      });
      await element.click();
      await delay(waitAfter);
      return true;
    } catch (e) {
      if (attempt === maxRetries) throw e;
      await delay(retryDelay);
    }
  }
};

// 금액 정제 함수
const cleanAmount = (amount) => (amount ? amount.replace(/[^0-9]/g, "") : "0");

// 팝업 데이터 추출 함수 (확장 버전)
const extractPopupData = async (page) => {
  const dataTemplate = {
    법원: "",
    사건번호: "",
    접수일자: "",
    사건명: "",
    문서명: "",
    소송등인지: "",
    송달료: "",
    법원보관금: "",
    납부금액: "",
    가상계좌번호: "",
  };

  try {
    await page.waitForSelector("#mf_pfwork_PSP2E1P03", {
      visible: true,
      timeout: 5000,
    });

    // 팝업 내 추가 데이터 추출
    const popupData = await page.evaluate(() => {
      const getText = (selector) =>
        document.querySelector(selector)?.textContent?.trim() || "";

      return {
        법원: getText('[data-title="법원"]'),
        사건번호: getText('[data-title="사건번호"]'),
        접수일자: getText('[data-title="가상계좌발급일자"]'),
        사건명: getText('tr:has([data-title="원고"]) td'),
        문서명: getText('tr:has([data-title="결제구분"]) td'),
        소송등인지: getText('[id*="grp_stmpAmt"] tr:first-child td'),
        송달료: getText('[id*="grp_dlvrf"] tr:first-child td'),
        법원보관금: getText('[id*="grp_cortCdpst"] tr:first-child td'),
        납부금액: getText('[data-title="납부금액"]'),
        가상계좌번호: getText('[data-title="가상계좌번호"]').replace(
          /[^0-9]/g,
          ""
        ),
      };
    });

    return {
      ...popupData,
      소송등인지: cleanAmount(popupData.소송등인지),
      송달료: cleanAmount(popupData.송달료),
      법원보관금: cleanAmount(popupData.법원보관금),
      납부금액: cleanAmount(popupData.납부금액),
    };
  } catch (e) {
    console.error("Popup Error:", e.message);
    return dataTemplate;
  }
};

// 행 데이터 추출 함수 (첫 행 특별 처리 추가)
const extractRowData = async (page, index, isFirstRow = false) => {
  const selector = `#mf_pfwork_grd_lwstSbmsnLst_body_table tr:not(.w2grid_header):not(.w2grid_emptyRow):nth-child(${
    index + 1
  })`;

  try {
    await page.waitForFunction(
      (sel) => {
        const row = document.querySelector(sel);
        return row && row.cells.length >= 5;
      },
      { timeout: 20000, polling: 500 },
      selector
    );

    let rowData = await page.evaluate((sel) => {
      const row = document.querySelector(sel);
      const cells = Array.from(row.cells);
      return {
        법원: cells[0]?.textContent?.trim(),
        접수일자: cells[1]?.textContent?.trim(),
        사건번호: cells[2]?.querySelector("a")?.textContent?.trim() || "",
        문서명: cells[3]?.querySelector("a")?.textContent?.trim() || "",
        사건명: cells[4]?.textContent?.trim(),
      };
    }, selector);

    // 첫 행이고 납부증 버튼이 있는 경우
    if (isFirstRow) {
      const payButtonSelector = `${selector} [data-col_id="payButton"] button`;
      const hasPaymentButton = await page
        .$(payButtonSelector)
        .catch(() => false);

      if (hasPaymentButton) {
        await safeClick(page, payButtonSelector, { waitAfter: 1000 });
        const popupData = await extractPopupData(page);
        await safeClick(page, "#mf_pfwork_PSP2E1P03_wframe_btn_close");
        await delay(500);

        // 팝업 데이터로 덮어쓰기
        rowData = {
          ...rowData,
          법원: popupData.법원 || rowData.법원,
          사건번호: popupData.사건번호 || rowData.사건번호,
          접수일자: popupData.접수일자 || rowData.접수일자,
          사건명: popupData.사건명 || rowData.사건명,
          문서명: popupData.문서명 || rowData.문서명,
        };
      }
    }

    return rowData;
  } catch (e) {
    console.error(`행 데이터 추출 실패 (인덱스: ${index})`, e.message);
    return null;
  }
};

// 세션 유지 함수
const handleSessionExtension = async (page) => {
  const extendSession = async () => {
    try {
      await page.waitForSelector("#mf_pfheader_btn_extTime:not([disabled])", {
        timeout: 10000,
        visible: true,
      });
      await safeClick(page, "#mf_pfheader_btn_extTime");
      console.log("🔄 세션 연장 버튼 클릭");

      await page.waitForSelector("#mf_pfheader_PSPLOGP02_wframe_btn_extend", {
        timeout: 10000,
        visible: true,
      });
      await safeClick(page, "#mf_pfheader_PSPLOGP02_wframe_btn_extend");
      console.log("✅ 세션 연장 완료");
      return true;
    } catch (e) {
      console.error("⚠️ 세션 연장 실패:", e.message);
      return false;
    }
  };

  const sessionInterval = setInterval(async () => {
    try {
      const success = await extendSession();
      if (!success) {
        console.log("🚨 세션 연장 실패, 재시도 중...");
        await delay(30000);
        await extendSession();
      }
    } catch (e) {
      console.error("🚨 세션 연장 최종 실패:", e.message);
    }
  }, 9 * 60 * 1000);

  return sessionInterval;
};

// 현재 페이지 번호 가져오기
const getCurrentPageNumber = async (page) => {
  return await page.evaluate(() => {
    const selectedPage = document.querySelector(".w2pageList_label_selected");
    return selectedPage ? parseInt(selectedPage.textContent) : 1;
  });
};

// 마지막 페이지 번호 가져오기
const getLastPageNumber = async (page) => {
  try {
    const totalItems = await page.evaluate(() => {
      const totalSpan = document.querySelector("#mf_pfwork_spn_total");
      return totalSpan ? parseInt(totalSpan.textContent.replace(/,/g, "")) : 0;
    });

    const itemsPerPage = await page.evaluate(() => {
      const perPageSelect = document.querySelector("#mf_pfwork_sbx_rows");
      if (!perPageSelect) return 30;
      const selectedOption = perPageSelect.options[perPageSelect.selectedIndex];
      const match = selectedOption.textContent.match(/(\d+)개씩/);
      return match ? parseInt(match[1]) : 30;
    });

    const lastPage = Math.ceil(totalItems / itemsPerPage);
    console.log(
      `총 항목: ${totalItems}, 페이지당 항목 수: ${itemsPerPage}, 마지막 페이지: ${lastPage}`
    );
    return lastPage;
  } catch (e) {
    console.error("마지막 페이지 번호 가져오기 실패:", e.message);
    return 999;
  }
};

// 특정 페이지로 이동
const goToPage = async (page, targetPage) => {
  try {
    const visiblePages = await page.evaluate(() => {
      const pageLinks = Array.from(
        document.querySelectorAll(".w2pageList_li_label a")
      );
      return pageLinks.map((link) => parseInt(link.textContent));
    });

    const minPage = Math.min(...visiblePages);
    const maxPage = Math.max(...visiblePages);

    if (targetPage < minPage) {
      await safeClick(page, "#mf_pfwork_pgl_tmprStrgLst_prev_btn", {
        waitAfter: 1500,
      });
      return await goToPage(page, targetPage);
    } else if (targetPage > maxPage) {
      await safeClick(page, "#mf_pfwork_pgl_tmprStrgLst_next_btn", {
        waitAfter: 1500,
      });
      return await goToPage(page, targetPage);
    }

    const pageSelector = `#mf_pfwork_pgl_tmprStrgLst_page_${targetPage}`;
    await safeClick(page, pageSelector, { waitAfter: 1500 });

    // 페이지 이동 후 데이터 로드 강화
    await page.waitForFunction(
      () => {
        const firstRow = document.querySelector(
          "#mf_pfwork_grd_lwstSbmsnLst_body_table tr:not(.w2grid_emptyRow):first-child"
        );
        return (
          firstRow &&
          firstRow.cells.length >= 5 &&
          firstRow.cells[0].textContent.trim() !== "" &&
          firstRow.cells[1].textContent.trim() !== ""
        );
      },
      { timeout: 30000 }
    );

    const currentPage = await getCurrentPageNumber(page);
    if (currentPage !== targetPage) {
      throw new Error(`페이지 이동 실패: ${currentPage} → ${targetPage}`);
    }

    return targetPage;
  } catch (e) {
    console.error(`페이지 ${targetPage}(으)로 이동 실패:`, e.message);
    throw e;
  }
};

// 다음 페이지로 이동
const goToNextPage = async (page, currentPage) => {
  const nextPage = currentPage + 1;
  return await goToPage(page, nextPage);
};

// 메인 실행 함수
(async () => {
  let browser;
  let sessionInterval;

  try {
    console.log("🔄 브라우저 연결 시도...");
    browser = await puppeteer.connect({
      browserURL: "http://localhost:9222",
      defaultViewport: null,
      ignoreHTTPSErrors: true,
    });
    console.log("✅ 브라우저 연결 성공");

    const pages = await browser.pages();
    const page = pages.find((p) => p.url().includes("ecfs.scourt.go.kr"));
    if (!page) throw new Error("대상 페이지를 찾을 수 없음");

    await page.setRequestInterception(true);
    page.on("request", (req) => {
      if (
        ["image", "stylesheet", "font", "media"].includes(req.resourceType())
      ) {
        req.abort();
      } else {
        req.continue();
      }
    });

    sessionInterval = await handleSessionExtension(page);

    const allData = [];
    let currentPage = await getCurrentPageNumber(page);
    const lastPage = await getLastPageNumber(page);

    while (currentPage <= lastPage) {
      console.log(
        `\n🔄 [${currentPage}페이지] 처리 시작 (마지막 페이지: ${lastPage})`
      );

      const rows = await page.$$(
        "#mf_pfwork_grd_lwstSbmsnLst_body_table tr:not(.w2grid_header):not(.w2grid_emptyRow)"
      );

      if (rows.length === 0) {
        console.log(`⚠️ ${currentPage}페이지에 데이터가 없습니다. 종료합니다.`);
        break;
      }

      for (let i = 0; i < rows.length; i++) {
        try {
          // 첫 행 여부 확인
          const isFirstRow = i === 0;
          const rowData = await extractRowData(page, i, isFirstRow);

          if (!rowData) {
            console.log(`⚠️ ${currentPage}-${i + 1} 행 무시`);
            continue;
          }

          // 필수 필드 검증 (법원명, 접수일자)
          if (!rowData.법원 || !rowData.접수일자) {
            console.log(
              `⚠️ ${currentPage}-${i + 1} 필수 필드 누락: 법원=${
                rowData.법원
              }, 접수일자=${rowData.접수일자}`
            );
            continue;
          }

          const payButtonSelector = `tr:nth-child(${
            i + 1
          }) [data-col_id="payButton"] button`;
          const hasPaymentButton = await page.$(payButtonSelector);

          let paymentData = {
            소송등인지: "",
            송달료: "",
            법원보관금: "",
            납부금액: "",
            가상계좌번호: "",
          };

          if (hasPaymentButton) {
            await safeClick(page, payButtonSelector, { waitAfter: 1000 });
            paymentData = await extractPopupData(page);
            await safeClick(page, "#mf_pfwork_PSP2E1P03_wframe_btn_close");
            await delay(500);
          }

          allData.push({
            ...rowData,
            ...paymentData,
            납부증여부: hasPaymentButton ? "있음" : "없음",
          });

          console.log(
            `✅ ${currentPage}-${i + 1} 완료 | 법원: ${
              rowData.법원
            } | 사건번호: ${rowData.사건번호} | 납부금액: ${
              paymentData.납부금액 || "없음"
            } | 가상계좌번호: ${paymentData.가상계좌번호 || "없음"} `
          );
        } catch (e) {
          console.error(`⚠️ ${currentPage}-${i + 1} 실패: ${e.message}`);
          await page.reload({ waitUntil: "networkidle2" });
          await delay(2000);
          continue;
        }
      }

      if (currentPage >= lastPage) break;

      try {
        currentPage = await goToNextPage(page, currentPage);
        console.log(`🔀 페이지 이동 완료: ${currentPage - 1} → ${currentPage}`);
      } catch (e) {
        console.error("페이지 이동 실패:", e.message);
        break;
      }
    }

    if (allData.length > 0) {
      await csvWriter.writeRecords(allData);
      console.log(
        `\n✅ CSV 저장 완료! 총 ${allData.length}건의 데이터가 저장되었습니다.`
      );

      // XLSX 변환
      const worksheet = xlsx.utils.json_to_sheet(allData);
      const workbook = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(workbook, worksheet, "사건상세");

      xlsx.writeFile(workbook, "법원_사건_상세_최종.xlsx");
      console.log(`✅ XLSX 저장 완료! "법원_사건_상세_최종.xlsx" 생성됨`);
    }
  } catch (error) {
    console.error("\n🚨 실행 중 오류 발생:", error);
    await fs.writeFile(
      `error-log-${Date.now()}.json`,
      JSON.stringify(error, Object.getOwnPropertyNames(error))
    );
  } finally {
    if (sessionInterval) clearInterval(sessionInterval);
    if (browser) await browser.disconnect();
    console.log("\n🏁 프로그램 종료");
  }
})();
