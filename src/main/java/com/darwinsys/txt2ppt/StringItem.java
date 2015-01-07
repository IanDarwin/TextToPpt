package com.darwinsys.txt2ppt;

public class StringItem extends Item {
	String item;

	/**
	 * @param item The text of the Item
	 */
	public StringItem(String item) {
		super();
		this.item = item;
	}

	@Override
	public String toString() {
		return item;
	}
	
}
