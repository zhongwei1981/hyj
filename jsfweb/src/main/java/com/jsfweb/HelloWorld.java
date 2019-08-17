package com.jsfweb;

import javax.faces.bean.ManagedBean;
import javax.faces.bean.SessionScoped;

@SessionScoped
@ManagedBean
public class HelloWorld implements java.io.Serializable {
	private static final long serialVersionUID = 1L;

	private String name = "Hello World";

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}
}
