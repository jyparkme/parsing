package onlinemed;

import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;

import java.time.LocalDate;
import java.util.HashMap;
import java.util.Map;
import java.util.Random;

public class ExcelBuild {

    private Map<Integer, String> idList = new HashMap<>();

    public void setIds() {
        idList.put(0, "본원");
        idList.put(1, "여의도");
        idList.put(2, "강남");
        idList.put(3, "수원");
        idList.put(4, "대구");
        idList.put(5, "부산");
        idList.put(6, "광주");
    }

    private String mainUrl = "http://hospital.onlinemed.co.kr/";
    private String loginUrl = mainUrl + "login.php";
    private String dataUrl = mainUrl + "reserve/down_rsv_list.php?scom=&yearid=2022&pip_yn=&sort1=&sort2=&dtestand=regdte&hpcchk=N&searchstart=2022-01-01&searchend=" + LocalDate.now().minusDays(1) + "&sword=&&dtestand=regdte";
    String fileName;
    Map<Integer, String[]> mList = new HashMap<>();

    public void setExcel(int seq) throws Exception{

        String id = idList.get(seq);

        // 로그인 정보
        Map<String, String> loginData = new HashMap<>();
        loginData.put("id", id);
        loginData.put("pw", id);
        String userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36 Edg/98.0.1108.62";

        // 메인쿠키 얻기
        Connection.Response tryLogin = Jsoup.connect(mainUrl)
//                .timeout(3000)
                .userAgent(userAgent)
                .header("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9")
                .header("Upgrade-Insecure-Requests", "1")
                .method(Connection.Method.GET)
                .execute();
        Map<String, String> tryCookie = tryLogin.cookies();

        // 메인쿠키로 로그인쿠키 얻기
        Connection.Response getLogin = Jsoup.connect(loginUrl)
                .userAgent(userAgent)
                .timeout(3000)
                .header("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9")
                .header("Content-Type", "application/x-www-form-urlencoded")
                .header("Upgrade-Insecure-Requests", "1")
                .header("sec-ch-ua", "\" Not A;Brand\";v=\"99\", \"Chromium\";v=\"99\", \"Microsoft Edge\";v=\"99\"")
                .header("sec-ch-ua-mobile", "?0")
                .header("sec-ch-ua-platform", "\"Windows\"")
                .cookies(tryCookie)
                .data(loginData)
                .method(Connection.Method.POST)
                .execute();
        Map<String, String> loginCookie = getLogin.cookies();

        // 로그인쿠키로 예약자리스트 요청
        Document doc = Jsoup.connect(dataUrl)
                .userAgent(userAgent)
                .header("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9")
                .header("Upgrade-Insecure-Requests", "1")
                .ignoreContentType(true)
                .cookies(loginCookie)
                .get();

        doc.selectFirst("tr").remove(); // 타이틀 제거
        Elements elements = doc.getElementsByTag("td");
        ExcelInfo info = new ExcelInfo();

        for (int i=0, k=0; i<elements.size(); i+=37, k++) {
            String[] temp = new String[117];
            for (int j=0; j<12; j++) {
                if (j == 10) {
                    StringBuilder sb = new StringBuilder();
                    String preSb = null;

                    for (int s : info.sIndex) {

                        if (elements.get(s+i).text() != null) {
                            sb.append(elements.get(s+i).text());
                        }
                    }
                    switch (elements.get(14+i).text()) {
                        case "0":
                            if (elements.get(15+i).text() != "Y") {
                                preSb = "당일" + (Integer.parseInt(elements.get(26 + i).text().replace(",", "")))/ 10000 + "만수납+공단미차감";
                                break;
                            } else {preSb = "당일" + (Integer.parseInt(elements.get(26+i).text().replace(",","")))/10000 + "만수납+공단차감";}
                            break;
                        case "50":
                            if (elements.get(15+i).text() != "Y") {
                                if (elements.get(26+i).text() != "0") {
                                    preSb = "당일" + (Integer.parseInt(elements.get(26+i).text().replace(",","")))/10000 + "만수납+공단미차감";
                                    break;
                                } else {preSb = "급공50%+공단미차감";
                                break;}
                            } else {
                                if (elements.get(26+i).text() != "0") {
                                    preSb = "당일" + (Integer.parseInt(elements.get(26+i).text().replace(",","")))/10000 + "만수납+공단차감";
                                    break;
                                } else {preSb = "급공50%+공단차감";
                                break;}
                            }
                        case "100":
                            if (elements.get(15+i).text() != "Y") {
                                preSb = "회사지원100%+공단미차감";
                                break;
                            } else {preSb = "회사지원100%+공단차감";
                            break;}
                    }

                    String reSb = preSb + new String(sb);
                    temp[10] = reSb;
                }else if(j == 11) {
                    Random rd = new Random();
                    switch (id) {
                        case "omca36":
                            temp[19] = info.c2[rd.nextInt(3)];
                        break;
                        case "omca34":
                            temp[19] = info.c3[rd.nextInt(4)];
                        break;
                        case "omca35":
                            temp[19] = info.c1[rd.nextInt(7)];
                        break;
                        case "omch33":
                            temp[19] = "24";
                        break;
                        case "omcf09":
                            temp[19] = "22";
                        break;
                        case "omcb05":
                            temp[19] = info.c6[rd.nextInt(2)];
                        break;
                        case "omcg23":
                            temp[19] = info.c7[rd.nextInt(2)];
                        break;
                    }
                }else{
                    temp[info.reIndex[j]] = elements.get((info.index[j]) + i).text();
                }
            }
            mList.put(k, temp);
        }

        switch (seq) {
            case 0: fileName = "광화문";
            break;
            case 1: fileName = "여의도";
            break;
            case 2: fileName = "강남";
            break;
            case 3: fileName = "수원";
            break;
            case 4: fileName = "대구";
            break;
            case 5: fileName = "부산";
            break;
            case 6: fileName = "광주";
            break;
        }

        System.out.println(fileName + "센터: 예약자명단 추출 완료");
    }
}
