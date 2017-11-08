package com.novayre.jidoka.robot.tutorial;

import java.net.URI;

import org.junit.Test;

public class EncodeUrlTest {

	@Test
	public void encodeURL() {
		URI uri = URI.create("http://blog.jidoka.io/image/ñúes_salvajes");
		System.out.println(uri.toASCIIString());
	}
}
