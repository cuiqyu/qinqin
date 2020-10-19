package com.littlebayo.project.qinq.pengda.controller;

import com.littlebayo.framework.web.controller.BaseController;
import com.littlebayo.framework.web.domain.AjaxResult;
import com.littlebayo.project.qinq.pengda.service.AttendanceService;
import org.apache.shiro.authz.annotation.RequiresPermissions;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;

/**
 * 芃达网络科技考勤报表管理
 *
 * @author cuiqiongyu
 * @date 2020-10-11 16:14
 */
@Controller
@RequestMapping("/qinq/pengda/attendance")
public class AttendanceController extends BaseController {

    private static final Logger logger = LoggerFactory.getLogger(AttendanceController.class);
    private String prefix = "qinq/pengda";

    @Autowired
    private AttendanceService attendanceService;

    /**
     * @param
     * @return java.lang.String
     * @description 考勤报表界面
     * @author cuiqiongyu
     * @date 20:57 2020-10-11
     **/
    @RequiresPermissions("qinq:pengda:attendance:view")
    @GetMapping()
    public String attendance() {
        return prefix + "/attendance";
    }

    /**
     * @param request
     * @return com.littlebayo.framework.web.domain.AjaxResult
     * @description 钉钉考勤文件上传
     * @author cuiqiongyu
     * @date 20:57 2020-10-11
     **/
    @PostMapping("/upload")
    @ResponseBody
    public AjaxResult uploadFile(HttpServletRequest request) throws Exception {
        try {
            // 转型为MultipartHttpRequest
            MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;
            // 获得文件
            MultipartFile file = multipartRequest.getFile("attendanceUpload");
            // 导入文件
            return attendanceService.importDingdingFile(file);
        } catch (Exception e) {
            return AjaxResult.error(e.getMessage());
        }
    }

}
