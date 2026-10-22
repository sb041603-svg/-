from pathlib import Path

code = r'''# ============================================================
# 충청도 26개 시군 엔트로피 분석용 전처리 통합 스크립트
# 입력: 충청도_데이터총합.zip 및 이미 추출된 충청도_데이터총합 폴더
# 출력: 충청도_엔트로피가중치_전처리_최종_R생성.xlsx
#
# 실행 방법:
# 1) RStudio에서 이 파일 전체를 Source 또는 Ctrl+Shift+Enter로 실행
# 2) 파일 선택창에서 충청도_데이터총합.zip 선택
# 3) 원자료 폴더는 기존에 확인한 경로를 사용
# ============================================================

# 0. 패키지 ----------------------------------------------------
필요패키지 <- c("readxl", "dplyr", "tidyr", "stringr", "purrr", "tibble", "writexl")

없는패키지 <- 필요패키지[
  !vapply(필요패키지, requireNamespace, logical(1), quietly = TRUE)
]
if (length(없는패키지) > 0) {
  install.packages(없는패키지)
}

suppressPackageStartupMessages({
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(stringr)
  library(purrr)
  library(tibble)
  library(writexl)
})

# 1. ZIP 및 원자료 폴더 ----------------------------------------
zip_path <- file.choose()  # 충청도_데이터총합.zip 선택

# 이전에 확인한 실제 원자료 폴더
work_dir <- "C:/Users/tm/Desktop/충청도/충청도_데이터총합"

if (!dir.exists(work_dir)) {
  # ZIP과 같은 폴더 안에 추출 폴더가 있는 경우도 확인
  후보폴더 <- file.path(dirname(zip_path), "충청도_데이터총합")
  if (dir.exists(후보폴더)) {
    work_dir <- 후보폴더
  } else {
    stop(
      "원자료 폴더를 찾지 못했습니다.\n",
      "확인한 경로: ", work_dir, "\n",
      "ZIP과 같은 폴더의 충청도_데이터총합 폴더도 없습니다."
    )
  }
}

파일목록 <- list.files(
  path = work_dir,
  recursive = TRUE,
  full.names = TRUE,
  pattern = "\\.xlsx$",
  ignore.case = TRUE
)

if (length(파일목록) == 0) {
  stop("원자료 폴더에 xlsx 파일이 없습니다: ", work_dir)
}

cat("원자료 폴더:", work_dir, "\n")
cat("발견한 xlsx 파일 수:", length(파일목록), "\n")
print(basename(파일목록))

# 파일 검색: 전체 경로가 아닌 파일명에 정규식을 적용
get_file <- function(pattern) {
  파일명 <- basename(파일목록)
  matched <- grepl(pattern, 파일명, perl = TRUE)
  found <- 파일목록[matched]

  if (length(found) != 1) {
    stop(
      "파일을 정확히 하나 찾지 못했습니다: ", pattern,
      "\n발견 수: ", length(found),
      "\n일치한 파일: ",
      if (length(found)) paste(basename(found), collapse = ", ") else "(없음)",
      "\n전체 파일 목록:\n",
      paste(파일명, collapse = "\n")
    )
  }
  found[[1]]
}

# 필수 파일 검색 테스트(파일이 없으면 여기서 중단)
pop_file <- get_file(
  "최종 정리본_7도_시군별_장애인_등록_자동차_인구_2020_2024.*xlsx$"
)
sat_file_chungnam <- get_file("교통수단_만족도_충남_2025_전처리완료V.*xlsx$")
sat_file_chungbuk <- get_file("교통수단_만족도_충북_2025_전처리완료V.*xlsx$")
gtfs_file <- get_file("충청도_서울직행_GTFS_연도통합분석.*xlsx$")

# 2. 분석지역: 충남 15개 + 충북 11개 --------------------------
regions <- bind_rows(
  tibble(
    시도 = "충청남도",
    시군 = c(
      "천안시", "공주시", "보령시", "아산시", "서산시",
      "논산시", "계룡시", "당진시", "금산군", "부여군",
      "서천군", "청양군", "홍성군", "예산군", "태안군"
    )
  ),
  tibble(
    시도 = "충청북도",
    시군 = c(
      "청주시", "충주시", "제천시", "보은군", "옥천군",
      "영동군", "증평군", "진천군", "괴산군", "음성군", "단양군"
    )
  )
)

stopifnot(nrow(regions) == 26, !anyDuplicated(regions[c("시도", "시군")]))

# 3. 2024 인구·장애인·자가용 자료 ----------------------------
pop_raw <- read_excel(pop_file, col_names = FALSE)

if (ncol(pop_raw) < 37) {
  stop("인구·장애인·자동차 원본의 열 개수가 37개 미만입니다.")
}

raw_sido <- trimws(as.character(pop_raw[[1]]))
raw_sigungu <- trimws(as.character(pop_raw[[2]]))

# 시도명이 병합 셀로 제공되는 경우 아래 행에 시도명을 이어 붙임
current_sido <- NA_character_
sido_vec <- rep(NA_character_, length(raw_sido))

for (i in seq_along(raw_sido)) {
  if (!is.na(raw_sido[i]) &&
      raw_sido[i] %in% c("충청남도", "충청북도")) {
    current_sido <- raw_sido[i]
  }
  sido_vec[i] <- current_sido
}

keep <- !is.na(sido_vec) &
  sido_vec %in% c("충청남도", "충청북도") &
  raw_sigungu %in% regions$시군

idx <- which(keep)

to_num <- function(x) {
  x <- trimws(as.character(x))
  x[x == ""] <- NA_character_
  suppressWarnings(as.numeric(gsub(",", "", x, fixed = TRUE)))
}

pop_selected <- tibble(
  시도 = sido_vec[idx],
  시군 = raw_sigungu[idx],
  전체장애인수 = to_num(pop_raw[[31]][idx]),
  전체자가용수 = to_num(pop_raw[[34]][idx]),
  `65세이상인구` = to_num(pop_raw[[36]][idx]),
  전체등록인구 = to_num(pop_raw[[37]][idx])
) %>%
  distinct(시도, 시군, .keep_all = TRUE)

pop <- regions %>%
  left_join(pop_selected, by = c("시도", "시군")) %>%
  mutate(
    노인비율_pct = if_else(
      !is.na(전체등록인구) & 전체등록인구 > 0,
      `65세이상인구` / 전체등록인구 * 100,
      NA_real_
    ),
    장애인비율_pct = if_else(
      !is.na(전체등록인구) & 전체등록인구 > 0,
      전체장애인수 / 전체등록인구 * 100,
      NA_real_
    ),
    자가용대수_인구당 = if_else(
      !is.na(전체등록인구) & 전체등록인구 > 0,
      전체자가용수 / 전체등록인구,
      NA_real_
    ),
    자가용대수_인구100명당 = 자가용대수_인구당 * 100
  )

cat("\n[인구·장애인·자동차 자료 확인]\n")
cat("분석 지역 수:", nrow(pop), "\n")
cat("장애인 수 결측:", sum(is.na(pop$전체장애인수)), "\n")
cat("자동차 수 결측:", sum(is.na(pop$전체자가용수)), "\n")
cat("노인 인구 결측:", sum(is.na(pop$`65세이상인구`)), "\n")
cat("등록인구 결측:", sum(is.na(pop$전체등록인구)), "\n")

# 4. 2025 순만족도 --------------------------------------------
read_sat <- function(f) {
  read_excel(f, sheet = "요약_불만족보통만족") %>%
    transmute(
      시도 = as.character(시도),
      시군 = as.character(시군),
      불만족비율 = suppressWarnings(as.numeric(`불만족비율(유효)`)),
      보통비율 = suppressWarnings(as.numeric(`보통비율(유효)`)),
      만족비율 = suppressWarnings(as.numeric(`만족비율(유효)`)),
      순만족도_pct = suppressWarnings(as.numeric(`순만족도(%)`))
    )
}

sat <- bind_rows(
  read_sat(sat_file_chungnam),
  read_sat(sat_file_chungbuk)
) %>%
  distinct(시도, 시군, .keep_all = TRUE)

# 5. 시외·고속·철도 원자료 지정 -------------------------------
source_specs <- tibble::tribble(
  ~시도,       ~교통수단, ~파일패턴,
  "충청남도",  "시외버스", "충남_시외_시군분류.*xlsx$",
  "충청남도",  "고속버스", "충남_고속_시군분류.*xlsx$",
  "충청남도",  "철도",     "충남_철도_시군분류.*xlsx$",
  "충청북도",  "시외버스", "충북_시외_시군분류.*xlsx$",
  "충청북도",  "고속버스", "충북_고속_시군분류.*xlsx$",
  "충청북도",  "철도",     "충북_철도_시군분류.*xlsx$"
)

read_transport <- function(시도, 교통수단, 파일패턴) {
  시도값 <- 시도
  수단값 <- 교통수단
  f <- get_file(파일패턴)

  d <- read_excel(f, sheet = "원자료_시군구추정") %>%
    mutate(
      연도 = suppressWarnings(as.integer(연도)),
      `월~일합계` = suppressWarnings(
        as.numeric(gsub(",", "", as.character(`월~일합계`), fixed = TRUE))
      ),
      `시/도` = as.character(`시/도`),
      시군구_추정 = as.character(시군구_추정)
    ) %>%
    filter(`시/도` == 시도값)

  data_out <- d %>%
    filter(시군구_추정 %in% regions$시군[regions$시도 == 시도값]) %>%
    group_by(
      시도 = `시/도`,
      시군 = 시군구_추정,
      연도
    ) %>%
    summarise(
      원자료합계 = if (all(is.na(`월~일합계`))) NA_real_ else
        sum(`월~일합계`, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      교통수단 = 수단값,
      원본파일 = basename(f)
    )

  unmatched_out <- d %>%
    filter(
      is.na(시군구_추정) |
        str_detect(coalesce(시군구_추정, ""), "미분류|확인")
    ) %>%
    transmute(
      원본파일 = basename(f),
      시도 = `시/도`,
      연도,
      교통수단 = 수단값,
      출발지,
      시군구_추정,
      분류근거,
      `월~일합계`
    )

  list(data = data_out, unmatched = unmatched_out)
}

read_list <- purrr::pmap(
  source_specs,
  function(시도, 교통수단, 파일패턴) {
    read_transport(시도, 교통수단, 파일패턴)
  }
)

transport_raw <- purrr::map_dfr(read_list, "data")
unmatched <- purrr::map_dfr(read_list, "unmatched")

if (nrow(transport_raw) == 0) {
  stop("transport_raw가 0행입니다. 원자료의 시도명·시군구 추정값을 확인하세요.")
}

cat("\n[교통 원자료 집계 행 수]", nrow(transport_raw), "\n")
print(transport_raw %>% count(시도, 교통수단, 연도), n = Inf)

# 6. 버스·철도 연도별 패널 및 증감률 ---------------------------
transport_yearly <- transport_raw %>%
  mutate(
    분류 = case_when(
      str_detect(교통수단, "시외|고속|버스") ~ "버스",
      str_detect(교통수단, "철도") ~ "철도",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(!is.na(분류)) %>%
  group_by(시도, 시군, 연도, 분류) %>%
  summarise(
    이용인원 = if (all(is.na(원자료합계))) NA_real_ else
      sum(원자료합계, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  pivot_wider(names_from = 분류, values_from = 이용인원)

transport_wide <- tidyr::crossing(regions, 연도 = 2021:2025) %>%
  left_join(transport_yearly, by = c("시도", "시군", "연도")) %>%
  arrange(시도, 시군, 연도) %>%
  group_by(시도, 시군) %>%
  mutate(
    버스이용인원증감률_pct = if_else(
      !is.na(버스) & !is.na(lag(버스)) & lag(버스) > 0,
      (버스 / lag(버스) - 1) * 100,
      NA_real_
    ),
    철도이용인원증감률_pct = if_else(
      !is.na(철도) & !is.na(lag(철도)) & lag(철도) > 0,
      (철도 / lag(철도) - 1) * 100,
      NA_real_
    )
  ) %>%
  ungroup()

stopifnot(nrow(transport_wide) == 26 * 5)

# 7. 서울 직행 GTFS 노선·운행회차 ------------------------------
gtfs_raw <- read_excel(gtfs_file, sheet = "연도별 지역요약") %>%
  mutate(분석연도 = suppressWarnings(as.integer(분석연도)))

필수_GTFS열 <- c(
  "도", "출발시군", "분석연도", "route_id기준_노선수",
  "직접서울행_운행회차수", "지역→서울_운행횟수", "서울→지역_운행횟수"
)
누락_GTFS열 <- setdiff(필수_GTFS열, names(gtfs_raw))
if (length(누락_GTFS열) > 0) {
  stop(
    "GTFS 파일에 필요한 열이 없습니다: ",
    paste(누락_GTFS열, collapse = ", ")
  )
}

gtfs_region <- gtfs_raw %>%
  group_by(시도 = 도, 시군 = 출발시군, 연도 = 분석연도) %>%
  summarise(
    서울직행노선수 = sum(`route_id기준_노선수`, na.rm = TRUE),
    서울직행운행횟수 = sum(직접서울행_운행회차수, na.rm = TRUE),
    지역서울행운행횟수 = sum(`지역→서울_운행횟수`, na.rm = TRUE),
    서울지역행운행횟수 = sum(`서울→지역_운행횟수`, na.rm = TRUE),
    .groups = "drop"
  )

gtfs_panel <- tidyr::crossing(regions, 연도 = 2021:2025) %>%
  left_join(gtfs_region, by = c("시도", "시군", "연도")) %>%
  mutate(
    GTFS상태 = if_else(
      is.na(서울직행노선수),
      "GTFS 요약 행 없음: 0 처리, 확인 필요",
      "GTFS 요약 집계"
    ),
    서울직행노선수 = replace_na(서울직행노선수, 0),
    서울직행운행횟수 = replace_na(서울직행운행횟수, 0),
    지역서울행운행횟수 = replace_na(지역서울행운행횟수, 0),
    서울지역행운행횟수 = replace_na(서울지역행운행횟수, 0)
  ) %>%
  arrange(시도, 시군, 연도) %>%
  group_by(시도, 시군) %>%
  mutate(
    서울직행노선수증감률_pct = if_else(
      lag(서울직행노선수) > 0,
      (서울직행노선수 / lag(서울직행노선수) - 1) * 100,
      NA_real_
    ),
    서울직행운행횟수증감률_pct = if_else(
      lag(서울직행운행횟수) > 0,
      (서울직행운행횟수 / lag(서울직행운행횟수) - 1) * 100,
      NA_real_
    )
  ) %>%
  ungroup()

stopifnot(nrow(gtfs_panel) == 26 * 5)

# 8. 최종 1행=1시군 테이블 ------------------------------------
final <- regions %>%
  left_join(pop, by = c("시도", "시군")) %>%
  left_join(sat, by = c("시도", "시군"))

# 버스·철도 이용원자료 증감률: 2022~2024년 각각 별도 열
for (yr in 2022:2024) {
  tmp <- transport_wide %>%
    filter(연도 == yr) %>%
    select(시도, 시군, 버스이용인원증감률_pct, 철도이용인원증감률_pct) %>%
    rename(
      !!paste0("버스이용인원증감률_", yr, "년_pct") :=
        버스이용인원증감률_pct,
      !!paste0("철도이용인원증감률_", yr, "년_pct") :=
        철도이용인원증감률_pct
    )
  final <- left_join(final, tmp, by = c("시도", "시군"))
}

# 서울 직행 노선수·운행횟수 증감률: 2022~2025년 각각 별도 열
for (yr in 2022:2025) {
  tmp <- gtfs_panel %>%
    filter(연도 == yr) %>%
    select(시도, 시군, 서울직행노선수증감률_pct, 서울직행운행횟수증감률_pct) %>%
    rename(
      !!paste0("서울직행노선수증감률_", yr, "년_pct") :=
        서울직행노선수증감률_pct,
      !!paste0("서울직행운행횟수증감률_", yr, "년_pct") :=
        서울직행운행횟수증감률_pct
    )
  final <- left_join(final, tmp, by = c("시도", "시군"))
}

# 2024 서울 직행 공급량을 인구 1만 명당으로 표준화
gtfs_2024 <- gtfs_panel %>%
  filter(연도 == 2024) %>%
  select(시도, 시군, 서울직행노선수, 서울직행운행횟수)

final <- final %>%
  left_join(gtfs_2024, by = c("시도", "시군")) %>%
  mutate(
    서울직행노선수_인구1만명당 = if_else(
      !is.na(전체등록인구) & 전체등록인구 > 0,
      서울직행노선수 / 전체등록인구 * 10000,
      NA_real_
    ),
    서울직행운행횟수_인구1만명당 = if_else(
      !is.na(전체등록인구) & 전체등록인구 > 0,
      서울직행운행횟수 / 전체등록인구 * 10000,
      NA_real_
    )
  )

# 엔트로피 입력용: 연도별 증감률을 합치지 않음
entropy_cols <- c(
  "시도", "시군", "노인비율_pct", "장애인비율_pct",
  "자가용대수_인구당", "순만족도_pct",
  paste0("버스이용인원증감률_", 2022:2024, "년_pct"),
  paste0("철도이용인원증감률_", 2022:2024, "년_pct"),
  paste0("서울직행노선수증감률_", 2022:2025, "년_pct"),
  paste0("서울직행운행횟수증감률_", 2022:2025, "년_pct"),
  "서울직행노선수_인구1만명당",
  "서울직행운행횟수_인구1만명당"
)
entropy_table <- final %>% select(any_of(entropy_cols))

# 9. 결측·계산불가 목록 ---------------------------------------
missing_table <- final %>%
  pivot_longer(
    cols = -c(시도, 시군),
    names_to = "항목",
    values_to = "값"
  ) %>%
  filter(is.na(값)) %>%
  mutate(
    상태 = "결측/계산 불가",
    사유 = "원자료 누락, 전년도 값 0, 또는 해당 연도 자료 부재 여부 확인"
  )

# 10. 지표 설명 ------------------------------------------------
definitions <- tibble(
  항목 = c(
    "분석 단위", "노인비율", "장애인비율", "차량 지표",
    "버스 이용 지표", "철도 이용 지표", "순만족도",
    "증감률 기준", "결측 처리", "엔트로피 전처리 주의"
  ),
  내용 = c(
    "충청남도 15개 시군 + 충청북도 11개 시군, 1행=1시군",
    "2024년 65세 이상 인구 / 전체 등록인구 × 100",
    "2024년 전체 장애인 수 / 전체 등록인구 × 100",
    "2024년 전체 자가용 수 / 전체 등록인구. 원자료 결측은 결측으로 유지하며 원본 확인 필요",
    "시외버스+고속버스의 원자료 월~일합계를 합산하고 전년 대비 증감률을 2022~2024년 별도 열로 보존",
    "철도 원자료 월~일합계를 합산하고 전년 대비 증감률을 2022~2024년 별도 열로 보존",
    "2025년 시군별 만족도 파일의 순만족도(%) 원값",
    "이용원자료는 2022~2024년, GTFS 공급지표는 2022~2025년 전년 대비 증감률을 각각 계산",
    "교통 원자료가 없는 지역·연도는 임의로 0으로 바꾸지 않고 결측 여부 확인. GTFS는 요약 행이 없는 경우 0 처리하므로 상태 열 확인",
    "증감률에는 음수가 포함되므로 엔트로피 계산 전 지표 방향 통일과 0 이상 정규화가 필요. 결측 대체는 근거를 남길 것"
  )
)

# 11. 저장 전 필수 검증 및 엑셀 저장 ---------------------------
stopifnot(
  nrow(regions) == 26,
  nrow(pop) == 26,
  nrow(final) == 26,
  nrow(transport_wide) == 130,
  nrow(gtfs_panel) == 130,
  exists("transport_raw"),
  exists("unmatched")
)

out_file <- file.path(
  dirname(zip_path),
  "충청도_엔트로피가중치_전처리_최종_R생성.xlsx"
)

write_xlsx(
  list(
    `01_엔트로피분석용_26개지역` = entropy_table,
    `02_전체지표_검토용` = final,
    `03_버스철도_이용원자료연도별` = transport_wide,
    `04_서울직행_GTFS연도별` = gtfs_panel,
    `05_교통수단합산원자료` = transport_raw,
    `06_2024_인구차량원자료` = pop,
    `07_2025_순만족도원자료` = sat,
    `08_미분류출발지점검` = unmatched,
    `09_결측및계산불가` = missing_table,
    `10_정의와전처리기준` = definitions
  ),
  path = out_file
)

if (!file.exists(out_file)) {
  stop("엑셀 저장 함수가 끝났지만 결과 파일을 찾을 수 없습니다: ", out_file)
}

cat("\n============================================\n")
cat("엑셀 저장 성공:\n", out_file, "\n")
cat("시군 수:", nrow(final), "\n")
cat("교통 원자료 행 수:", nrow(transport_raw), "\n")
cat("버스·철도 패널 행 수:", nrow(transport_wide), "\n")
cat("GTFS 패널 행 수:", nrow(gtfs_panel), "\n")
cat("============================================\n")
'''
out='/mnt/data/충청도_엔트로피_전처리_통합코드.R'
Path(out).write_text(code, encoding='utf-8')
print(f"생성 완료: {out}\n코드 줄 수: {len(code.splitlines())}")
