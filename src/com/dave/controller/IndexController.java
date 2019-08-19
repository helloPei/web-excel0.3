package com.dave.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

/**
 * @Author: Dave
 * @Date: 2019/8/8 14:59
 * @Description: TODO
 */
@Controller
@RequestMapping("/")
public class IndexController {
    /**
     * 主页面跳转
     *
     * @return index.jsp
     */
    @RequestMapping("doIndexUI")
    public String doIndexUI() {
        return "index";
    }
}
