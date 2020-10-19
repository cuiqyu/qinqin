package com.littlebayo.project.qinq.pengda.service;

import com.littlebayo.framework.web.domain.AjaxResult;
import org.springframework.web.multipart.MultipartFile;

/**
 * 考勤服务
 *
 * @author cuiqiongyu
 * @date 2020/10/19 14:52
 */
public interface AttendanceService {

    /**
     * 导入钉钉文件
     *
     * @param file 上传的文件
     * @return 返回结果
     */
    AjaxResult importDingdingFile(MultipartFile file) throws Exception;

}
