package com.littlebayo.project.qinq.pengda.controller;

import com.littlebayo.common.utils.file.FileUploadUtils;
import com.littlebayo.framework.config.RuoYiConfig;
import com.littlebayo.framework.web.controller.BaseController;
import com.littlebayo.framework.web.domain.AjaxResult;
import org.apache.shiro.authz.annotation.RequiresPermissions;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

/**
 * 芃达网络科技考勤报表管理
 *
 * @author cuiqiongyu
 * @date 2020-10-11 16:14
 */
@Controller
@RequestMapping("/qinq/pengda/attendance")
public class AttendanceController extends BaseController {

    private String prefix = "qinq/pengda";

    private static final Logger logger = LoggerFactory.getLogger(AttendanceController.class);

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
     * @param file
     * @return com.littlebayo.framework.web.domain.AjaxResult
     * @description 钉钉考勤文件上传
     * @author cuiqiongyu
     * @date 20:57 2020-10-11
     **/
    @PostMapping("/upload")
    @ResponseBody
    public AjaxResult uploadFile(MultipartFile file) throws Exception {
        try {
            // 参数校验
            if (null == file || file.getSize() <= 0) {
                logger.error("处理文件失败，上传的文件大小不能为空。");
                return AjaxResult.error("上传的文件大小不能为空！");
            }
            // 文件后缀必须是.xls,.xlsx的
            String fileName = file.getName();
            if (!fileName.endsWith(".xls") && !fileName.endsWith("xlsx")) {
                logger.error("处理文件失败，上传的文件类型不正确，文件类型只能是.xls或.xlsx。当前文件类型：{}",
                        fileName.substring(fileName.lastIndexOf(".")));
                return AjaxResult.error("上传的文件类型不正确，文件类型只能是.xls或.xlsx！");
            }

            return AjaxResult.success();
        } catch (Exception e) {
            return AjaxResult.error(e.getMessage());
        }
    }

}
