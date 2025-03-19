import {createContext} from 'react';

const testContext = createContext();
const Test = (props) => {
    return (
        <testContext.Provider value={props}>
            {props.children}
        </testContext.Provider>
    )
}

export default Test;