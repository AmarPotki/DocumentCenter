import { ITaxonomyPickerProps } from "react-taxonomypicker";
export interface ITaxonomyPickerLoaderProps extends ITaxonomyPickerProps {
    onRender?: (domElement: HTMLElement, context?: any) => void;
    onDispose?: (domElement: HTMLElement, context?: any) => void;
    context?: any;
}