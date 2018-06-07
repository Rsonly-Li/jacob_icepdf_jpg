import java.io.Serializable;
import java.util.*;

/**
 * Created by 李綦睿 on 2017-05-05 0005.
 */
public class MyProperties extends Properties implements Serializable {

    private static final long serialVersionUID = -241081438420794870L;

    private final LinkedHashSet<Object> keys = new LinkedHashSet<Object>();

    @Override
    public synchronized Object put(Object key,Object value){
        keys.add(key);
        return super.put(key,value);
    }

    @Override
    public Set<String> stringPropertyNames(){
        Set<String> set = new LinkedHashSet<String>();
        for (Object obj : keys){
            set.add((String) obj);
        }
        return set;
    }

    @Override
    public Set<Object> keySet(){
        return this.keys;
    }

    public Enumeration<Object> keys(){
        return Collections.<Object>enumeration(keys);
    }

    private String url = getClass().getResource("libra.properties").getPath();

    public String getUrl() {
        return url;
    }
}
