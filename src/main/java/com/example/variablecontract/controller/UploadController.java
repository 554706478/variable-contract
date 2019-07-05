package com.example.variablecontract.controller;


import com.example.variablecontract.model.TStudent;
import com.example.variablecontract.service.TStudentService;
import com.github.pagehelper.PageHelper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import java.util.List;

/**
 * <p>
 *  前端控制器
 * </p>
 *
 * @author v-wuchengs
 * @since 2019-07-01
 */
@Controller
@RequestMapping("/tStudent")
public class UploadController {

    @Autowired
    TStudentService tStudentService;

    @RequestMapping("getPage")
    @ResponseBody
    public List<TStudent> getByPage() {
        PageHelper.startPage(1, 3);//页数 每页数据数
        return tStudentService.getByPage();
    }

}

