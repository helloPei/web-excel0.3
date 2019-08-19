package com.dave.common.vo;

import java.io.Serializable;
/**
 * JSON结果集实体对象
 * @author davewpw
 *
 */
public class JsonResult implements Serializable{
	private static final long serialVersionUID = -2040132524942880840L;
	private int state = 0;//error
	private String message = "ok";//返回的message
	private Object data;//返回的数据
	
	public JsonResult() {}
	public JsonResult(String message) {
		this.message = message;
	}
	public JsonResult(String message, int state) {
		this.state = state;
		this.message = message;
	}
	public JsonResult(Object data){
		this.state = 1;//succeed
		this.data = data;
	}
	public JsonResult(Throwable e){
		this.message = e.getMessage();
	}
	public int getState() {
		return state;
	}
	public void setState(int state) {
		this.state = state;
	}
	public String getMessage() {
		return message;
	}
	public void setMessage(String message) {
		this.message = message;
	}
	public Object getData() {
		return data;
	}
	public void setData(Object data) {
		this.data = data;
	}
}