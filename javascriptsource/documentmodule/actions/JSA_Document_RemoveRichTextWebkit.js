// This file was generated by Mendix Studio Pro.
//
// WARNING: Only the following code will be retained when actions are regenerated:
// - the import list
// - the code between BEGIN USER CODE and END USER CODE
// - the code between BEGIN EXTRA CODE and END EXTRA CODE
// Other code you write will be lost the next time you deploy the project.
import "mx-global";
import { Big } from "big.js";

// BEGIN EXTRA CODE
// END EXTRA CODE

/**
 * @returns {Promise.<void>}
 */
export async function JSA_Document_RemoveRichTextWebkit() {
	// BEGIN USER CODE
	
	//500ms 뒤에 실행 되도록 SetTimeOut 함수 사용(Page Load 후 Webkit을 삭제하기 위해)
	setTimeout(() =>{

		let TooltipElement = document.querySelectorAll('.cke_browser_webkit');
		let RoofCount = TooltipElement.length;
		
		for(var i = 0 ; i< RoofCount; i++){
			TooltipElement[i].classList.add('RomoveWebkit')
		}
	}, 500);

	// END USER CODE
}