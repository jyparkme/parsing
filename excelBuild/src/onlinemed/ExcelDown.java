package onlinemed;

import org.apache.http.Header;
import org.apache.http.NameValuePair;
import org.apache.http.client.config.RequestConfig;
import org.apache.http.client.entity.UrlEncodedFormEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.message.BasicNameValuePair;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.net.URI;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class ExcelDown {

    public String id;
    public int idNum;
    public String url = "http://hospital.onlinemed.co.kr/";
    public String loginUrl = url+"login.php";
    public String downUrl = url+"reserve/down_rsv_list.php?scom=&yearid=2022&pip_yn=&sort1=&sort2=&dtestand=regdte&hpcchk=N&searchstart=2022-01-01&searchend="+ LocalDate.now().minusDays(1)+"&sword=&&dtestand=regdte";

    public String findId(int idNum) {
        List<String> ids = new ArrayList<>();
        ids.add("omca36");
        ids.add("omca34");
        ids.add("omca35");
        ids.add("omch33");
        ids.add("omcf09");
        ids.add("omcb05");
        ids.add("omcg23");

        return ids.get(idNum);
    }

    public void getFileDown() throws Exception {
        URI uri = new URI(loginUrl);
        RequestConfig config = RequestConfig.custom().setSocketTimeout(10000).setConnectTimeout(10000).setConnectionRequestTimeout(10000).build();
        CloseableHttpClient httpClient = HttpClients.createDefault();
        HttpPost httpPost = new HttpPost(loginUrl);
        httpPost.setConfig(config);

        List<NameValuePair> data = new ArrayList<>();
        data.add(new BasicNameValuePair("id", id));
        data.add(new BasicNameValuePair("pw", id));
        httpPost.setEntity(new UrlEncodedFormEntity(data, "euc_kr"));
        CloseableHttpResponse response = httpClient.execute(httpPost);
        Header[] headers = response.getHeaders("Set-Cookie");
        for (Header headerElem : headers) {
            String cookies = "";
            cookies = headerElem.getValue();
        }

        String fileName = "";
        switch (idNum) {
            case 0:
                fileName = "(원본)광화문.xls";
                break;
            case 1:
                fileName = "(원본)여의도.xls";
                break;
            case 2:
                fileName = "(원본)강남.xls";
                break;
            case 3:
                fileName = "(원본)수원.xls";
                break;
            case 4:
                fileName = "(원본)대구.xls";
                break;
            case 5:
                fileName = "(원본)부산.xls";
                break;
            case 6:
                fileName = "(원본)광주.xls";
                break;
        }
        String filePath = "C:\\Onlinemed\\원본(" + LocalDate.now().minusDays(1) + ")";
        HttpGet httpGet = new HttpGet(downUrl);
        CloseableHttpResponse responseFile = httpClient.execute(httpGet);
        BufferedInputStream bInStr = new BufferedInputStream(responseFile.getEntity().getContent());
        BufferedOutputStream bOutStr = new BufferedOutputStream(new FileOutputStream(new File(filePath + "\\" + fileName)));
        int inpByte;
        while ((inpByte = bInStr.read()) != -1) {
            bOutStr.write(inpByte);
        }
        bInStr.close();
        bOutStr.close();
        response.close();
        System.out.println(fileName+" Complete");
    }
}