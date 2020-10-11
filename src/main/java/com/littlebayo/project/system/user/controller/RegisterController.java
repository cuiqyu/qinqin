package com.littlebayo.project.system.user.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import com.littlebayo.framework.shiro.service.RegisterService;
import com.littlebayo.framework.web.controller.BaseController;
import com.littlebayo.framework.web.domain.AjaxResult;
import com.littlebayo.project.system.config.service.IConfigService;
import com.littlebayo.project.system.user.domain.User;

/**
 * 注册验证
 * 
 * @author littlebayo
 */
@Controller
public class RegisterController extends BaseController
{
    @Autowired
    private RegisterService registerService;

    @Autowired
    private IConfigService configService;

    @GetMapping("/register")
    public String register()
    {
        return "register";
    }

    @PostMapping("/register")
    @ResponseBody
    public AjaxResult ajaxRegister(User user)
    {
        if (!("true".equals(configService.selectConfigByKey("sys.account.registerUser"))))
        {
            return error("当前系统没有开启注册功能！");
        }
        String msg = registerService.register(user);
        return StringUtils.isEmpty(msg) ? success() : error(msg);
    }
}
