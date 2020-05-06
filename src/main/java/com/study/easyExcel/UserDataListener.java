package com.study.easyExcel;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.study.dao.UserDao;
import com.study.pojo.User;

import java.util.ArrayList;
import java.util.List;

public class UserDataListener extends AnalysisEventListener<User> {

    private static final int BATCH_COUNT = 3000;
    List<User> list = new ArrayList<User>();

    private UserDao userDao;

    public UserDataListener(){
        userDao= new UserDao();
    }

    public UserDataListener(UserDao userDao){
        this.userDao = userDao;
    }


    //解析数据
    @Override
    public void invoke(User user, AnalysisContext analysisContext) {
        System.out.println(user);
        list.add(user);
        if(list.size()>=BATCH_COUNT){
            saveUser();
            list.clear();
        }
    }

    private void saveUser() {
        //执行插入操作
        userDao.save(list);
    }

    //数据解析完之后调用
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
       //如果最后一次解析不足3000条 剩下的数据还需要执行插入语句
       saveUser();
    }
}
