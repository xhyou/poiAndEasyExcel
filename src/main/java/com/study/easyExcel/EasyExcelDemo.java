package com.study.easyExcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.study.pojo.User;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.ForkJoinPool;
import java.util.concurrent.RecursiveTask;

public class EasyExcelDemo {

    final  static String PATH= "D:\\dev_workhome\\studyPoi\\test\\";

    /**
     * EasyExcel的写操作
     */
    @Test
    public void easyExcelWrite() throws IOException {
        ForkJoinPool forkJoinPool = new ForkJoinPool();
        ListUtils listUtils = new ListUtils(0,1000);
        List<User>  listUsers = (List<User>) forkJoinPool.invoke(listUtils);
//        EasyExcel.write(PATH+"easyExcelTest.xlsx", User.class).sheet("sheet1").doWrite(listUsers);
        //写法二 一般用在重复多次写入一个Excel中
        String fileName = PATH+ "easyExcelTest" + System.currentTimeMillis() + ".xlsx";
          ExcelWriter excelWriter = EasyExcel.write(fileName, User.class).build();
          WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
          for (int i = 0; i <2; i++) {
            // 分页去数据库查询数据 这里可以去数据库查询每一页的数据
            List<User> data = listUsers;//data() 数据库中查询的数据;
            excelWriter.write(data, writeSheet);
          }
          excelWriter.finish();
    }

    @Test
    public void easyExcelRead(){
        String fileName = PATH+"easyExcelTest.xlsx";
        EasyExcel.read(fileName,User.class,new UserDataListener()).sheet().doRead();
    }

    class ListUtils extends RecursiveTask {
        List<User> userList = new ArrayList<>();
        static final int MAX_NUM = 500;

        int start, end;

        ListUtils(int s, int e) {
            start = s;
            end = e;
        }

        @Override
        protected List<User> compute() {
            if(end-start <= MAX_NUM) {
                for(int i=start; i<end; i++){
                    User user = createUser(i);
                    userList.add(user);
                }
                return userList;
            } else {
                int middle = start + (end-start)/2;
                ListUtils subTask1 = new ListUtils(start, middle);
                ListUtils subTask2 = new ListUtils(middle, end);
                subTask1.fork();
                subTask2.fork();
                List<User> subTask1List = (List<User>)subTask1.join();
                List<User> subTask2List = (List<User>)subTask2.join();
                subTask1List.addAll(subTask2List);
                return  subTask1List;
            }
        }

        public User createUser(int i){
            User user = new User();
            user.setName("小名"+i);
            user.setSex(true);
            user.setDoubleData(Double.valueOf(i));
            user.setDate(new Date());
            user.setAge(i);
            return user;
        }
    }
}
