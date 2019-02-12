package com.luxiaoyuan.excel.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class LogUtils {

	/**
     * 获取controller日志logger
     *
     * @return
     */
    public static Logger getControllerLogger() {
        return LoggerFactory.getLogger(LogEnum.CONTROLLER.getCategory());
    }
 
    /**
     * 获取service日志logger
     *
     * @return
     */
    public static Logger getServiceLogger() {
        return LoggerFactory.getLogger(LogEnum.SERVICE.getCategory());
    }

    /**
     * 获取utils日志logger
     *
     * @return
     */
    public static Logger getUtilsLogger(){
        return LoggerFactory.getLogger(LogEnum.UTILS.getCategory());
    }

}
