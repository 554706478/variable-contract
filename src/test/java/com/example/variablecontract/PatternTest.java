package com.example.variablecontract;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PatternTest {

    private static final Logger log = LoggerFactory.getLogger(PatternTest.class);


    @Test
    public void pattern() {
        List<String> list = new ArrayList<>();

        String reg2="\\{\\w+\\}";
//        String reg1="\\$\\{\\w+\\}";

        String source="尊敬的${EmpName}，您好，\n" +
                "\n" +
                "截止到{currentDate}日，您尚有异常考勤信息未确认，未保证您的考勤数据完整性及准确性，请务必尽快在IPSA中完成{currentMonth}月份缺勤补录，谢谢！";
        Pattern pattern=Pattern.compile(reg2);
        Matcher matcher=pattern.matcher(source);
        while(matcher.find()){
            int start=matcher.start();//匹配到结果在源字符串的起始索引
            int end=matcher.end();//匹配到结果在源字符串的结束索引
            String group=matcher.group();//匹配到结果字符串
            list.add(group);
            System.out.println("group:"+group+",start="+start+",end="+end);
        }
        System.out.println(list.toString());
    }




}
