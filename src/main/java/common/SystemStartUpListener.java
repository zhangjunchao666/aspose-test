package common;

import javax.servlet.ServletContext;
import javax.servlet.ServletContextEvent;
import javax.servlet.ServletContextListener;
import javax.servlet.annotation.WebListener;

//服务器启动或停止(卸载)时执行监听器。
//将监听器的全类名,放入web.xml的<listener>标签中即可完成配置
@WebListener
public class SystemStartUpListener implements ServletContextListener {
	
	//服务器启动时要干的活
	@Override
	public void contextInitialized(ServletContextEvent sce) {
		//获取ApplicationServletContext对象
		ServletContext application = sce.getServletContext();
		//获取上下文路径getContextPath,设置到application的自定义中(APP_PATH)
		application.setAttribute("APP_PATH", application.getContextPath());

    }

	//服务器停止(卸载)时要干的活
	@Override
	public void contextDestroyed(ServletContextEvent sce) {
		
	}

}
