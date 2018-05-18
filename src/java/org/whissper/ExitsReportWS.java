package org.whissper;

import javax.jws.WebService;
import javax.jws.WebMethod;
import javax.jws.WebParam;
import javax.jws.soap.SOAPBinding;
import javax.xml.ws.Holder;

/**
 * ConsumptionWS
 * @author SAV2
 */
@WebService(serviceName = "ExitsReportWS")
@SOAPBinding(parameterStyle = SOAPBinding.ParameterStyle.WRAPPED)
public class ExitsReportWS {

    @WebMethod(operationName = "loadXLSX")
    public void loadXLSX(@WebParam(name = "startDate", mode = WebParam.Mode.IN) String startDateValue,
                         @WebParam(name = "endDate", mode = WebParam.Mode.IN) String endDateValue,
                         @WebParam(name = "depID", mode = WebParam.Mode.IN) String depIDValue,
                         @WebParam(name = "reference", mode = WebParam.Mode.OUT) Holder<String> refValue)
    {
        refValue.value = new ExcelLoaderEngine("C:/SAV2/nodeserver/exitlog/frontend/getfile/", startDateValue, endDateValue, depIDValue).loadData();
    }
    
}
