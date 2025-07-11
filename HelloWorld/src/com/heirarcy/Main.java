package com.heirarcy;

public class Main {
	public static void main(String[] args) {
		admin a = new admin();
        a.login();
        a.managewebsite();
        System.out.println();

        saller s = new saller();
        s.login();
        s.addProduct();
        System.out.println();

        customer c = new customer();
        c.login();
        c.browsePrduct();
	}
}
