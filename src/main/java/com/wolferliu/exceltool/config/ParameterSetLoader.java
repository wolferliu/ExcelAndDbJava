package com.wolferliu.exceltool.config;

import com.thoughtworks.xstream.XStream;
import com.thoughtworks.xstream.XStreamer;

import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.security.Policy;

/**
 * Created by IntelliJ IDEA.
 * Date: 3/14/11
 *
 * @author liubing
 */
public class ParameterSetLoader {
	XStream xs=new XStream();

	public ParameterSetLoader() {
		xs.processAnnotations(ParameterSet.class);
		xs.processAnnotations(ConnectionParameter.class);
//		xs.alias("ParameterSet", ParameterSet.class);
//		xs.alias("Connection", ConnectionParameter.class);
	}

	public ParameterSet loadConfig() throws IOException, ClassNotFoundException {
		ParameterSet ps=new ParameterSet();
		if(getConfigFile().exists()){
			ps= (ParameterSet) xs.fromXML(new FileReader(getConfigFile()));
		}
		return ps;
	}

	public void saveConfig(ParameterSet ps) throws IOException {
		xs.toXML(ps, new FileWriter(getConfigFile()));
	}

	private File configFile=null;
	private File getConfigFile(){
		if(configFile==null){
			configFile=new File(new File(System.getProperty("user.dir")), "config.xml");
		}
		return configFile;
	}

}
