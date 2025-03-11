package com.idea.controller;

import com.idea.DatasourceAnalyzer;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

/**
 * @ClassName: HelloController
 * @Description: TODO <br>
 * @Author: yuanbao
 * @Date: 2025/3/10
 **/
@RestController
public class HelloController
{
    @Autowired
    private DatasourceAnalyzer datasourceAnalyzer;

    @RequestMapping("/")
    public String getHello()
    {
        return "hello23";
    }

    @RequestMapping(value = "/hello", method = RequestMethod.GET)
    @ResponseBody
    public String sayHello()
    {
        return "Hello, World!";
    }

    @RequestMapping(value = "/t1", method = RequestMethod.GET)
    @ResponseBody
    public String solveV55Data()
    {
        try
        {
            datasourceAnalyzer.getDataFromMysql();
            return "solveV55Data finished!";

        } catch (IOException e)
        {
            e.printStackTrace();
            return "solveV55Data failed!";
        }
    }

}