package com.artmoa;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.SocketTimeoutException;
import java.net.URL;
import java.net.URLEncoder;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class NaverExhibitionCollector {

    // =========================
    // 0. 네이버 API 설정
    // =========================
    private static final String CLIENT_ID = "Q2rhzi7EM0YEl7RnbVFW";
    private static final String CLIENT_SECRET = "aeKTJ4jaVM";

    private static final ObjectMapper objectMapper = new ObjectMapper();

    // 필요하면 고정 날짜로 테스트 가능
    // private static final LocalDate TODAY = LocalDate.of(2026, 4, 21);
    private static final LocalDate TODAY = LocalDate.now();

    // =========================
    // Rate Limit 대응 설정
    // =========================
    private static final int REQUEST_DELAY_MS = 500;   // 모든 API 호출 전 기본 대기
    private static final int MAX_RETRIES = 5;          // 최대 재시도 횟수
    private static final long BASE_BACKOFF_MS = 1000L; // 1초, 2초, 4초, 8초, 16초

    // =========================
    // Excel 셀 길이 제한
    // =========================
    private static final int EXCEL_CELL_MAX_LENGTH = 32767;
    private static final String TRUNCATED_SUFFIX = "... [TRUNCATED]";

    public static void main(String[] args) {
        List<String> queries = buildQueries();
        System.out.println("생성된 검색어 수: " + queries.size());

        LinkedHashSet<String> candidateUrls = new LinkedHashSet<>();

        for (String query : queries) {
            try {
                System.out.println("검색어 처리 중: " + query);

                candidateUrls.addAll(searchNaverAllPages("blog", query));
                candidateUrls.addAll(searchNaverAllPages("cafearticle", query));
                candidateUrls.addAll(searchNaverAllPages("webkr", query));
                candidateUrls.addAll(searchNaverAllPages("news", query));

            } catch (Exception e) {
                System.out.println("검색 API 호출 실패: " + query);
                e.printStackTrace();
            }
        }

        System.out.println("후보 URL 수집 수: " + candidateUrls.size());

        List<Exhibition> exhibitions = new ArrayList<>();
        Set<String> dedupKeys = new HashSet<>();

        for (String url : candidateUrls) {
            try {
                if (!isAllowedUrl(url)) {
                    System.out.println("URL 제외: " + url);
                    continue;
                }

                System.out.println("상세 파싱 중: " + url);

                String html = fetchHtml(url);
                Exhibition exhibition = parseExhibition(html, url);

                if (!isValid(exhibition)) {
                    System.out.println("저장 제외: " + url);
                    continue;
                }

                String dedupKey = buildDedupKey(exhibition);
                if (dedupKeys.contains(dedupKey)) {
                    System.out.println("중복 제외: " + exhibition.title + " / " + url);
                    continue;
                }

                dedupKeys.add(dedupKey);
                exhibitions.add(exhibition);

            } catch (Exception e) {
                System.out.println("상세 파싱 실패: " + url);
            }
        }

        System.out.println("최종 저장 대상 수: " + exhibitions.size());

        try {
            writeExcel(exhibitions, "output/exhibitions.xlsx");
            System.out.println("엑셀 저장 완료: output/exhibitions.xlsx");
        } catch (Exception e) {
            System.out.println("엑셀 저장 실패");
            e.printStackTrace();
        }
    }

    // =========================
    // 1. 검색어 생성
    // =========================
    private static List<String> buildQueries() {
        List<String> regions = Arrays.asList(
                // 경기도 전체 시/군
                "경기", "경기도",
                "수원", "성남", "의정부", "안양", "부천", "광명", "평택", "동두천",
                "안산", "고양", "과천", "구리", "남양주", "오산", "시흥", "군포",
                "의왕", "하남", "용인", "파주", "이천", "안성", "김포", "화성",
                "광주", "양주", "포천", "여주",
                "연천", "가평", "양평",

                // 인천 전체 구/군
                "인천", "인천광역시",
                "중구", "동구", "미추홀구", "연수구", "남동구", "부평구", "계양구", "서구",
                "강화군", "옹진군"
        );

        List<String> artTerms = Arrays.asList(
                "미술 전시", "미술전", "전시회", "전시", "개인전", "기획전",
                "초대전", "그룹전", "회화 전시", "사진 전시", "조각 전시",
                "청년작가 전시", "아트 전시", "갤러리 전시", "미술관 전시",
                "작품 전시", "예술 전시", "미술 작품 전시", "전시 일정"
        );

        LinkedHashSet<String> queries = new LinkedHashSet<>();

        for (String region : regions) {
            for (String term : artTerms) {
                queries.add(region + " " + term);
            }
        }

        return new ArrayList<>(queries);
    }

    // =========================
    // 2. 네이버 검색 API 호출
    // =========================
    private static List<String> searchNaverAllPages(String type, String query) throws Exception {
        List<String> allUrls = new ArrayList<>();

        // start는 1, 31, 61, 91 순으로 조회
        for (int start = 1; start <= 91; start += 30) {
            try {
                allUrls.addAll(searchNaver(type, query, 30, start));
            } catch (Exception e) {
                System.out.println("페이지 검색 실패 -> type=" + type + ", query=" + query + ", start=" + start);
                throw e;
            }
        }

        return allUrls;
    }

    private static List<String> searchNaver(String type, String query, int display, int start) throws Exception {
        int attempt = 0;

        while (true) {
            try {
                Thread.sleep(REQUEST_DELAY_MS);

                String encodedQuery = URLEncoder.encode(query, "UTF-8");

                String apiUrl = "https://openapi.naver.com/v1/search/" + type
                        + ".json?query=" + encodedQuery
                        + "&display=" + display
                        + "&start=" + start
                        + "&sort=date";

                HttpURLConnection conn = (HttpURLConnection) new URL(apiUrl).openConnection();
                conn.setRequestMethod("GET");
                conn.setRequestProperty("X-Naver-Client-Id", CLIENT_ID);
                conn.setRequestProperty("X-Naver-Client-Secret", CLIENT_SECRET);

                int responseCode = conn.getResponseCode();
                BufferedReader br;

                if (responseCode == 200) {
                    br = new BufferedReader(new InputStreamReader(conn.getInputStream()));
                } else {
                    br = new BufferedReader(new InputStreamReader(conn.getErrorStream()));
                }

                StringBuilder response = new StringBuilder();
                String line;
                while ((line = br.readLine()) != null) {
                    response.append(line);
                }
                br.close();

                String responseBody = response.toString();

                if (responseCode == 200) {
                    JsonNode root = objectMapper.readTree(responseBody);
                    JsonNode items = root.get("items");

                    List<String> urls = new ArrayList<>();
                    if (items != null && items.isArray()) {
                        for (JsonNode item : items) {
                            String link = getText(item, "link");
                            if (!isBlank(link)) {
                                urls.add(cleanLink(link));
                            }
                        }
                    }

                    return urls;
                }

                boolean isRateLimited =
                        responseCode == 429 ||
                                responseBody.contains("\"errorCode\":\"012\"") ||
                                responseBody.contains("Rate limit exceeded") ||
                                responseBody.contains("속도 제한을 초과");

                if (isRateLimited && attempt < MAX_RETRIES) {
                    long backoffMs = BASE_BACKOFF_MS * (1L << attempt);
                    System.out.println("속도 제한 감지 -> " + backoffMs + "ms 대기 후 재시도"
                            + " | type=" + type + ", query=" + query + ", start=" + start + ", attempt=" + (attempt + 1));
                    Thread.sleep(backoffMs);
                    attempt++;
                    continue;
                }

                throw new RuntimeException("네이버 API 오류(" + type + "): " + responseBody);

            } catch (SocketTimeoutException e) {
                if (attempt < MAX_RETRIES) {
                    long backoffMs = BASE_BACKOFF_MS * (1L << attempt);
                    System.out.println("타임아웃 -> " + backoffMs + "ms 대기 후 재시도"
                            + " | type=" + type + ", query=" + query + ", start=" + start + ", attempt=" + (attempt + 1));
                    Thread.sleep(backoffMs);
                    attempt++;
                } else {
                    throw e;
                }
            }
        }
    }

    private static String getText(JsonNode node, String field) {
        JsonNode value = node.get(field);
        return value == null ? "" : value.asText();
    }

    private static String cleanLink(String link) {
        if (link == null) return "";
        return link.replace("&amp;", "&").trim();
    }

    // =========================
    // 3. 상세 HTML 가져오기
    // =========================
    private static String fetchHtml(String url) throws Exception {
        Document doc = Jsoup.connect(url)
                .userAgent("Mozilla/5.0")
                .timeout(10000)
                .followRedirects(true)
                .get();

        return doc.html();
    }

    // =========================
    // 4. 전시 정보 파싱
    // =========================
    private static Exhibition parseExhibition(String html, String url) {
        Document doc = Jsoup.parse(html);
        String fullText = preprocessText(doc.text());

        Exhibition exhibition = new Exhibition();
        exhibition.sourceUrl = defaultIfBlank(url, "없음");
        exhibition.sourceType = defaultIfBlank(detectSourceType(url), "UNKNOWN");

        exhibition.title = defaultIfBlank(extractTitle(doc), "없음");
        fillDates(fullText, exhibition);
        exhibition.locationName = defaultIfBlank(extractLocation(fullText), "없음");
        exhibition.address = defaultIfBlank(extractAddress(fullText), "없음");
        exhibition.region = defaultIfBlank(detectRegion(fullText, url, exhibition.locationName, exhibition.address), "UNKNOWN");
        exhibition.status = defaultIfBlank(detectStatus(exhibition.startDate, exhibition.endDate), "UNKNOWN");

        return exhibition;
    }

    private static String preprocessText(String text) {
        if (text == null) return "";
        return text.replaceAll("\\(월\\)|\\(화\\)|\\(수\\)|\\(목\\)|\\(금\\)|\\(토\\)|\\(일\\)", "")
                .replaceAll("\\s+", " ")
                .trim();
    }

    private static String extractTitle(Document doc) {
        Element ogTitle = doc.selectFirst("meta[property=og:title]");
        if (ogTitle != null && !ogTitle.attr("content").isBlank()) {
            return cleanTitle(ogTitle.attr("content"));
        }

        if (doc.title() != null && !doc.title().isBlank()) {
            return cleanTitle(doc.title());
        }

        Element h1 = doc.selectFirst("h1");
        if (h1 != null && !h1.text().isBlank()) {
            return cleanTitle(h1.text());
        }

        Element h2 = doc.selectFirst("h2");
        if (h2 != null && !h2.text().isBlank()) {
            return cleanTitle(h2.text());
        }

        return null;
    }

    private static String cleanTitle(String title) {
        if (title == null) return null;
        return title.replaceAll("<[^>]*>", "")
                .replaceAll("\\s+", " ")
                .trim();
    }

    // =========================
    // 5. 날짜 파싱
    // =========================
    private static void fillDates(String text, Exhibition exhibition) {
        if (matchFullDotDate(text, exhibition)) return;
        if (matchFullKorDate(text, exhibition)) return;
        if (matchSameYearShortEndDot(text, exhibition)) return;
        if (matchSameYearShortEndKor(text, exhibition)) return;
        if (matchDateRangeSlash(text, exhibition)) return;
        if (matchLooseDotDate(text, exhibition)) return;
        if (matchMonthDayOnly(text, exhibition)) return;
    }

    // 2026.04.01 ~ 2026.04.30 / 2026-04-01 ~ 2026-04-30
    private static boolean matchFullDotDate(String text, Exhibition exhibition) {
        Pattern pattern = Pattern.compile(
                "(20\\d{2})[.-]\\s*(\\d{1,2})[.-]\\s*(\\d{1,2})\\s*[~\\-–]\\s*(20\\d{2})[.-]\\s*(\\d{1,2})[.-]\\s*(\\d{1,2})"
        );
        Matcher m = pattern.matcher(text);

        if (m.find()) {
            exhibition.startDate = LocalDate.of(
                    Integer.parseInt(m.group(1)),
                    Integer.parseInt(m.group(2)),
                    Integer.parseInt(m.group(3))
            );
            exhibition.endDate = LocalDate.of(
                    Integer.parseInt(m.group(4)),
                    Integer.parseInt(m.group(5)),
                    Integer.parseInt(m.group(6))
            );
            return true;
        }
        return false;
    }

    // 2026/04/01 ~ 2026/04/30
    private static boolean matchDateRangeSlash(String text, Exhibition exhibition) {
        Pattern pattern = Pattern.compile(
                "(20\\d{2})\\s*/\\s*(\\d{1,2})\\s*/\\s*(\\d{1,2})\\s*[~\\-–]\\s*(20\\d{2})\\s*/\\s*(\\d{1,2})\\s*/\\s*(\\d{1,2})"
        );
        Matcher m = pattern.matcher(text);

        if (m.find()) {
            exhibition.startDate = LocalDate.of(
                    Integer.parseInt(m.group(1)),
                    Integer.parseInt(m.group(2)),
                    Integer.parseInt(m.group(3))
            );
            exhibition.endDate = LocalDate.of(
                    Integer.parseInt(m.group(4)),
                    Integer.parseInt(m.group(5)),
                    Integer.parseInt(m.group(6))
            );
            return true;
        }
        return false;
    }

    // 2026년 4월 1일 ~ 2026년 4월 30일
    private static boolean matchFullKorDate(String text, Exhibition exhibition) {
        Pattern pattern = Pattern.compile(
                "(20\\d{2})\\s*년\\s*(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일\\s*[~\\-–]\\s*(20\\d{2})\\s*년\\s*(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일"
        );
        Matcher m = pattern.matcher(text);

        if (m.find()) {
            exhibition.startDate = LocalDate.of(
                    Integer.parseInt(m.group(1)),
                    Integer.parseInt(m.group(2)),
                    Integer.parseInt(m.group(3))
            );
            exhibition.endDate = LocalDate.of(
                    Integer.parseInt(m.group(4)),
                    Integer.parseInt(m.group(5)),
                    Integer.parseInt(m.group(6))
            );
            return true;
        }
        return false;
    }

    // 2026.04.01 ~ 04.30
    private static boolean matchSameYearShortEndDot(String text, Exhibition exhibition) {
        Pattern pattern = Pattern.compile(
                "(20\\d{2})[.-]\\s*(\\d{1,2})[.-]\\s*(\\d{1,2})\\s*[~\\-–]\\s*(\\d{1,2})[.-]\\s*(\\d{1,2})"
        );
        Matcher m = pattern.matcher(text);

        if (m.find()) {
            int year = Integer.parseInt(m.group(1));
            exhibition.startDate = LocalDate.of(
                    year,
                    Integer.parseInt(m.group(2)),
                    Integer.parseInt(m.group(3))
            );
            exhibition.endDate = LocalDate.of(
                    year,
                    Integer.parseInt(m.group(4)),
                    Integer.parseInt(m.group(5))
            );
            return true;
        }
        return false;
    }

    // 2026년 4월 1일 ~ 4월 30일
    private static boolean matchSameYearShortEndKor(String text, Exhibition exhibition) {
        Pattern pattern = Pattern.compile(
                "(20\\d{2})\\s*년\\s*(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일\\s*[~\\-–]\\s*(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일"
        );
        Matcher m = pattern.matcher(text);

        if (m.find()) {
            int year = Integer.parseInt(m.group(1));
            exhibition.startDate = LocalDate.of(
                    year,
                    Integer.parseInt(m.group(2)),
                    Integer.parseInt(m.group(3))
            );
            exhibition.endDate = LocalDate.of(
                    year,
                    Integer.parseInt(m.group(4)),
                    Integer.parseInt(m.group(5))
            );
            return true;
        }
        return false;
    }

    // 2026. 4. 1. ~ 2026. 5. 10.
    private static boolean matchLooseDotDate(String text, Exhibition exhibition) {
        Pattern pattern = Pattern.compile(
                "(20\\d{2})\\.\\s*(\\d{1,2})\\.\\s*(\\d{1,2})\\.\\s*[~\\-–]\\s*(20\\d{2})\\.\\s*(\\d{1,2})\\.\\s*(\\d{1,2})\\."
        );
        Matcher m = pattern.matcher(text);

        if (m.find()) {
            exhibition.startDate = LocalDate.of(
                    Integer.parseInt(m.group(1)),
                    Integer.parseInt(m.group(2)),
                    Integer.parseInt(m.group(3))
            );
            exhibition.endDate = LocalDate.of(
                    Integer.parseInt(m.group(4)),
                    Integer.parseInt(m.group(5)),
                    Integer.parseInt(m.group(6))
            );
            return true;
        }
        return false;
    }

    // 4월 1일 ~ 5월 10일
    private static boolean matchMonthDayOnly(String text, Exhibition exhibition) {
        Pattern pattern = Pattern.compile(
                "(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일\\s*[~\\-–]\\s*(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일"
        );
        Matcher m = pattern.matcher(text);

        if (m.find()) {
            int year = TODAY.getYear();
            exhibition.startDate = LocalDate.of(
                    year,
                    Integer.parseInt(m.group(1)),
                    Integer.parseInt(m.group(2))
            );
            exhibition.endDate = LocalDate.of(
                    year,
                    Integer.parseInt(m.group(3)),
                    Integer.parseInt(m.group(4))
            );
            return true;
        }
        return false;
    }

    // =========================
    // 6. 장소 / 주소 파싱
    // =========================
    private static String extractLocation(String text) {
        Pattern pattern = Pattern.compile(
                "([가-힣A-Za-z0-9\\s]+?(갤러리|미술관|아트센터|아트홀|문화센터|문화예술회관|플랫폼|전시장|문화공간|복합문화공간|아트스페이스|스튜디오|창작공간|대안공간|프로젝트룸|쇼룸|카페|북카페|살롱|공간|센터))"
        );
        Matcher m = pattern.matcher(text);

        if (m.find()) {
            return m.group(1).replaceAll("\\s+", " ").trim();
        }
        return null;
    }

    private static String extractAddress(String text) {
        Pattern pattern = Pattern.compile(
                "((경기도|인천광역시|인천|경기)\\s+[가-힣A-Za-z0-9\\s\\-]+)"
        );
        Matcher m = pattern.matcher(text);

        if (m.find()) {
            return m.group(1).replaceAll("\\s+", " ").trim();
        }
        return null;
    }

    // =========================
    // 7. 검증 / 필터링
    // =========================
    private static boolean isValid(Exhibition exhibition) {
        if (exhibition == null) return false;
        return !isBlank(exhibition.sourceUrl);
    }

    private static boolean isAllowedUrl(String url) {
        if (isBlank(url)) return false;

        String lower = url.toLowerCase();

        if (lower.contains("adcr.naver.com")) return false;
        if (lower.endsWith(".jpg") || lower.endsWith(".png") || lower.endsWith(".gif")) return false;

        return lower.startsWith("http://") || lower.startsWith("https://");
    }

    private static String buildDedupKey(Exhibition e) {
        return normalize(e.sourceUrl);
    }

    private static String normalize(String value) {
        if (value == null) return "";
        return value.replaceAll("\\s+", "").toLowerCase();
    }

    // =========================
    // 8. 출처 / 지역 추정
    // =========================
    private static String detectSourceType(String url) {
        String lower = url.toLowerCase();

        if (lower.contains("blog")) return "BLOG";
        if (lower.contains("cafe")) return "CAFE";
        if (lower.contains("news")) return "NEWS";
        return "WEB";
    }

    private static String detectRegion(String text, String url, String locationName, String address) {
        String addressBase = safe(address).toLowerCase();
        String textBase = (safe(text) + " " + safe(url)).toLowerCase();
        String locationBase = safe(locationName).toLowerCase();

        if (containsIncheonKeyword(addressBase)) return "INCHEON";
        if (containsGyeonggiKeyword(addressBase)) return "GYEONGGI";

        if (containsIncheonKeyword(textBase)) return "INCHEON";
        if (containsGyeonggiKeyword(textBase)) return "GYEONGGI";

        if (containsIncheonKeyword(locationBase)) return "INCHEON";
        if (containsGyeonggiKeyword(locationBase)) return "GYEONGGI";

        return "UNKNOWN";
    }

    private static boolean containsIncheonKeyword(String text) {
        return text.contains("인천")
                || text.contains("인천광역시")
                || text.contains("중구")
                || text.contains("동구")
                || text.contains("미추홀")
                || text.contains("미추홀구")
                || text.contains("연수구")
                || text.contains("남동구")
                || text.contains("부평구")
                || text.contains("계양구")
                || text.contains("서구")
                || text.contains("강화군")
                || text.contains("옹진군");
    }

    private static boolean containsGyeonggiKeyword(String text) {
        return text.contains("경기")
                || text.contains("경기도")
                || text.contains("수원")
                || text.contains("성남")
                || text.contains("의정부")
                || text.contains("안양")
                || text.contains("부천")
                || text.contains("광명")
                || text.contains("평택")
                || text.contains("동두천")
                || text.contains("안산")
                || text.contains("고양")
                || text.contains("과천")
                || text.contains("구리")
                || text.contains("남양주")
                || text.contains("오산")
                || text.contains("시흥")
                || text.contains("군포")
                || text.contains("의왕")
                || text.contains("하남")
                || text.contains("용인")
                || text.contains("파주")
                || text.contains("이천")
                || text.contains("안성")
                || text.contains("김포")
                || text.contains("화성")
                || text.contains("광주")
                || text.contains("양주")
                || text.contains("포천")
                || text.contains("여주")
                || text.contains("연천")
                || text.contains("가평")
                || text.contains("양평");
    }

    // =========================
    // 9. 상태 판별
    // =========================
    private static String detectStatus(LocalDate startDate, LocalDate endDate) {
        if (startDate == null || endDate == null) {
            return "UNKNOWN";
        }

        if (endDate.isBefore(TODAY)) {
            return "ENDED";
        }

        if ((startDate.isBefore(TODAY) || startDate.isEqual(TODAY))
                && (endDate.isAfter(TODAY) || endDate.isEqual(TODAY))) {
            return "ONGOING";
        }

        if (startDate.isAfter(TODAY)) {
            return "UPCOMING";
        }

        return "UNKNOWN";
    }

    // =========================
    // 10. 엑셀 저장
    // =========================
    private static void writeExcel(List<Exhibition> exhibitions, String filePath) throws Exception {
        File outputFile = new File(filePath);
        File parent = outputFile.getParentFile();

        if (parent != null && !parent.exists()) {
            parent.mkdirs();
        }

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Exhibitions");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("전시명");
        header.createCell(1).setCellValue("시작일");
        header.createCell(2).setCellValue("종료일");
        header.createCell(3).setCellValue("장소명");
        header.createCell(4).setCellValue("주소");
        header.createCell(5).setCellValue("지역");
        header.createCell(6).setCellValue("출처유형");
        header.createCell(7).setCellValue("출처URL");
        header.createCell(8).setCellValue("상태");

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

        int rowNum = 1;
        for (Exhibition e : exhibitions) {
            Row row = sheet.createRow(rowNum);

            row.createCell(0).setCellValue(toExcelSafeText("title", e.title, rowNum));
            row.createCell(1).setCellValue(toExcelSafeText(
                    "startDate",
                    e.startDate != null ? e.startDate.format(formatter) : "없음",
                    rowNum
            ));
            row.createCell(2).setCellValue(toExcelSafeText(
                    "endDate",
                    e.endDate != null ? e.endDate.format(formatter) : "없음",
                    rowNum
            ));
            row.createCell(3).setCellValue(toExcelSafeText("locationName", e.locationName, rowNum));
            row.createCell(4).setCellValue(toExcelSafeText("address", e.address, rowNum));
            row.createCell(5).setCellValue(toExcelSafeText("region", defaultIfBlank(e.region, "UNKNOWN"), rowNum));
            row.createCell(6).setCellValue(toExcelSafeText("sourceType", defaultIfBlank(e.sourceType, "UNKNOWN"), rowNum));
            row.createCell(7).setCellValue(toExcelSafeText("sourceUrl", e.sourceUrl, rowNum));
            row.createCell(8).setCellValue(toExcelSafeText("status", defaultIfBlank(e.status, "UNKNOWN"), rowNum));

            rowNum++;
        }

        // autoSizeColumn 제거
        // 폭은 고정값으로만 설정
        sheet.setColumnWidth(0, 12000); // 전시명
        sheet.setColumnWidth(1, 4000);  // 시작일
        sheet.setColumnWidth(2, 4000);  // 종료일
        sheet.setColumnWidth(3, 10000); // 장소명
        sheet.setColumnWidth(4, 12000); // 주소
        sheet.setColumnWidth(5, 5000);  // 지역
        sheet.setColumnWidth(6, 5000);  // 출처유형
        sheet.setColumnWidth(7, 15000); // 출처URL
        sheet.setColumnWidth(8, 5000);  // 상태

        try (FileOutputStream fos = new FileOutputStream(outputFile)) {
            workbook.write(fos);
        } finally {
            workbook.close();
        }
    }

    // =========================
    // 11. 유틸
    // =========================
    private static boolean isBlank(String value) {
        return value == null || value.trim().isEmpty();
    }

    private static String safe(String value) {
        return value == null ? "" : value;
    }

    private static String defaultIfBlank(String value, String defaultValue) {
        return isBlank(value) ? defaultValue : value;
    }

    private static String toExcelSafeText(String fieldName, String value, int rowNum) {
        String text = defaultIfBlank(value, "없음");

        if (text.length() > EXCEL_CELL_MAX_LENGTH) {
            System.out.println("[WARN] Excel 셀 길이 초과 -> row=" + rowNum
                    + ", field=" + fieldName
                    + ", originalLength=" + text.length());

            int maxTextLength = EXCEL_CELL_MAX_LENGTH - TRUNCATED_SUFFIX.length();
            if (maxTextLength < 0) {
                maxTextLength = EXCEL_CELL_MAX_LENGTH;
            }

            return text.substring(0, maxTextLength) + TRUNCATED_SUFFIX;
        }

        return text;
    }

    // =========================
    // 12. 내부 클래스
    // =========================
    static class Exhibition {
        String title;
        LocalDate startDate;
        LocalDate endDate;
        String locationName;
        String address;
        String sourceUrl;
        String sourceType;
        String region;
        String status;
    }
}