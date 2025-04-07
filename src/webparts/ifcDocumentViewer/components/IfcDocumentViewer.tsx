import * as React from "react";
import { useState, useEffect } from "react";
import styles from "./IfcDocumentViewer.module.scss";
import { IDocumentLibraryViewerProps } from "./IIfcDocumentViewerProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {
  CommandBar,
  ICommandBarItemProps,
} from "@fluentui/react/lib/CommandBar";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
} from "@fluentui/react/lib/DetailsList";
import { Text } from "@fluentui/react/lib/Text";
import { IconButton } from "@fluentui/react/lib/Button";
import { IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";

export interface IDocument {
  Id: string;
  Title: string;
  Name: string;
  ModifiedBy: string;
  Modified: string;
  FileRef: string;
  FileType: string;
}

export interface IDocumentLibraryViewerState {
  documents: IDocument[];
  isLoading: boolean;
  error: string | null;
}

const DocumentLibraryViewer: React.FC<IDocumentLibraryViewerProps> = (
  props
) => {
  const [state, setState] = useState<IDocumentLibraryViewerState>({
    documents: [],
    isLoading: true,
    error: null,
  });

  const getDocuments = (): void => {
    const { context, libraryName } = props;
    const libraryToUse = libraryName || "Documents"; // Default to Documents if not specified

    const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${libraryToUse}')/items?$select=Id,Title,FileLeafRef,Modified,FileRef,Editor/Title&$expand=Editor&$orderby=Modified desc`;

    context.spHttpClient
      .get(endpoint, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          throw new Error(`Error fetching documents: ${response.statusText}`);
        }
      })
      .then((data) => {
        if (data && data.value) {
          const docs = data.value.map((item: any) => ({
            Id: item.Id,
            Title: item.Title || "No Title",
            Name: item.FileLeafRef,
            ModifiedBy: item.Editor ? item.Editor.Title : "Unknown",
            Modified: new Date(item.Modified).toLocaleDateString(),
            FileRef: item.FileRef,
            FileType: item.FileLeafRef.split(".").pop().toLowerCase(),
          }));

          setState({
            documents: docs,
            isLoading: false,
            error: null,
          });
        }
      })
      .catch((error) => {
        setState({
          documents: [],
          isLoading: false,
          error: error.message,
        });
        console.error("Error fetching documents:", error);
      });
  };

  useEffect(() => {
    getDocuments();
  }, []);

  const openDocument = (documentUrl: string): void => {
    window.open(documentUrl, "_blank");
  };

  const onLoadIFC = (): void => {
    // Implement your custom action here
    alert(
      "Custom action triggered! You can implement your specific functionality here."
    );
    // For example: Open a modal, trigger a flow, etc.
  };

  const commandItems: ICommandBarItemProps[] = [
    {
      key: "refresh",
      text: "Refresh",
      iconProps: { iconName: "Refresh" },
      onClick: getDocuments,
    },
    {
      key: "loadIFC",
      text: "Load to IFC Viewer",
      onRenderIcon: () => {
        return (
          <img
            src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAM8AAADPCAMAAABlX3VtAAAABGdBTUEAALGPC/xhBQAAAAFzUkdCAK7OHOkAAAMAUExURf///////enh0QBWobcokwCfs+cASACjt+cAUABcpbUkkQBao+cATgChtbMgjwBUn7Umk7UikbMej+cASgBYo7UmkRRoq/Fml8HX6QBYoecATDS1xeskaOP19+kMWP/9/QCjtfX7/cvt8b9An/Fsm//9/wVepaPd5fHV61jDz+3L5fvT4fn9/eGn07Mcjff5/cFIo7w4m3Snzfm3zTqBuU6/zfSFqwalt+soagOjt9WFwucARvPd7t2bzPvZ5eGl0lCPwanf592dzR6tv6nJ4fz9//3x9v3//wBUoQOlt92Zy3anzf3v9Pm1zWfJ0/f5+1CRweH19xutve5Efr/X6fXf76vJ4fFol7gtlYez1MHp7enx9wVep/v9/fHX6wdgp+fx91PBzomz1dN+v/vX5dWHwxSrvDK1xPf9/R6tvYax01rD0QBSnwBco9+dzUy/zQ+pu+vj0w5kqe1AfPFum+kKVuPt9e9GgcFGo8pgsEm+y6PF3/72+hRmq9nx9FKRwfvJ2hlsrfvW4s3t8RSpu+sobP/7/afG4O74+snr7+ssbfP5+wCdshJnq/eoxLcqlfWVtzZ+uCexwjy5yCBxsNfm8e01dUmMvt+fz13G0fN0odHh72efybrm7L0/n7jR5bEcjfjn8ukYYbs0mcre7fN+p+W12+W12dF2uwqmuTyDur9CoJvb49/z9o+413mq0Oe22/zi7bLO5Kbf5t7q8+kPWdWCwO/R58dYq7Uhj//5/OnA4MNMpdeKxO9Kg/WGrQChs+30+cPp7vFunecCUPesxsxotOfv9/m+02vK1fLp0+vE4uf3+f3p8PixySZ1s+Wx2OcCUnPN13Cly9mQx43W35a92qnh5za3xTZ8t6LD3TR8t1uWxesgZvegvtuUyfvQ34Kv073V5/3d6HzQ2vnr9fPZ681sta/h58/t8e9RieXv9aPF3TR+t6Hd5Oi73fWNsd2dz/FklzR7tfBbj8VWq/Xh781ut8HZ6dHv8a/h6dN8va3h58darfe1yza1xUWEWc4AABb/SURBVHgB7Z0LkBXVmcfP7Nw7NTV37h1g6qJ3IMywK6iIVaAo4MD62IRbpYjlQiIhhVPGXcA41IIaFZYl4BpX5cogIqCCXEXKEMXEKD4QzSTlO7JGjals1hUxPsotNVubSDalm+z3nUf3Od3n1d13dMuyh7n9mL7d59f/7/ufr8/tGQj5XE3dnysaQr7g+f8t6Bf6fKHPp3kFPut4m/XGv77+XxteP/GqfY2h/kx5Zq///bRvd7bA1LV381HPNgLpM+SZ/eNpnb2d7a106mrpbd/8enaiBvAsvO3XZ722Zctp27614/IEQfPUtN5OxsJf21s6566pJjiCbtfMPA8e+qCvo6Of/qus/c7XztGdJb5t1WXtLQoNrLz0ZO+YDfFdE23JyHP5KyM7OgpLxVTpr/zuH33OP/uJXh5oKlRX560+bzfvk43n/bX9FcHC5x2V0x40n47/ZPbcXpUjWGtvucz5btsOWXgmvlJpjtDgav/aHbYzws9mz43FWqOAMvCc8x0p0mSuSt+XrUA2nNbWbAql5wEcGUJetgPZcTICpeY55zQjztKlNiAXTjagtDwWdVApM5AbJxNQSh4HDgCN1+eQD04WoHQ81mBjmaQH8sNBoKOslmL8YSoeDxy9Qr446RVKw+MMNq5QzLb9cVIrlILHSx2dKSTBSQuUnMcbJ+pyyXBShlxiHs9gi4dcUpx0QEl5EqhDQy6w7eQ4qUIuIU9CnNDl0uCkAUrGkxhHAN1/qbmiDkprzULifigRT6LcYRnETeF1/e2bhiCyKSlQEp4U6gjb7j5ZDHxE2utcTQiUgCclDlfIAtTZZaNKBuTPkxqH59DJXdoRg9bWlkunWZMrEZA3TwYcO1DLFauuGt0wIF8eF07FcncHSUSrba1CLVfMaiLHNQzIk8eF0zFlWyUcthLWJs3pDZ4GCHCw9m8YkB+PE+ffLycnNHsARV2u5YrZ7FamUUBePD44JA1QL1OngQr58Dhx/o2NW7sUwlpOsW0ebI1UyIPHifOBGIZ3AeENnpRDCo5PDo1j5LZXN48/jjvkqEKiH4rgNAbIyZMEJxFQDMcHyCYN/ZmLx4nDc0ecyDvkJCsQ720EkIPHiQNGrU6eQBp18Dhu23bkkJ3HiRNRB5vkBWTA8QHCc5gnK48T58aoOngiD6Dq+gOmJrkVMr2TbrfxOHGmPKM9tgsoGFPQvtsNZAs5C09aHLdCWYEsjTb/KD2OB9C/aKXhG90Kmd9t5EmNc+C845xAhawKGUPOxJMaBz7qnfZ3TqBKX0aFTEAGngw4La0tPkBZFTKEnJ4nNc4s+sn1pwKkb7l2a2qcpv9kzxW0fBdC7izHDV5WhbRN121MjUN+3PkkG3qiQM5+KGMO6UJOw5Me540xwZCUD1Bh/I90TRLb3LatMYU4T3qc7ie6uDwgkg9Q5YbbRON18zRAMZ70OOTZl+RxTh9TaP5gnQ5EbHMDiT2DeZQnAw75vfo8mw9Qx6GgJboFN1A05CI8WXDuvzPIHskU7C5XGGl/GMsNFAFQV7PgkKdiw+oeOdT8A50u4TYnUOR5OYUnEw65VQ031MgNVNj0k7DxuiUXUKf6RKPMkw2HzI3p4wNU2aGjkLY5gLrGvCHtLP8+RvU1+5h6h+H2jR9u1fUaHqaQLYeaX5Gbo1t2AHXOvUZ6l6TP13UPG4ZD6g4ccuC7ETvgpuAoTitTpCZIDZMWHUBdT0n7hge7q886nu7CIQf+Q8vDqm2zQoUXvyE1R79oB+q8VBIo4PnK96NPtobSwJITh8yepos3NAVUyAhUGHmzHkLeagd6aU24b8Dzo2zqELJvs4HHcT90ibXm4S21foInP+MseKpbbNnjVgfO+4SJx65Qnw+PdaCxSwo4wXNbnxJe6ooXDtkV73+YJVhDzive4GpZQq79zvuDgBM891nk8cMhJ4bFteAI5uYcKmzSD+IFLRQLFiApgQTPFrMbeOKQ2QaDo1BGoMqNC0WLHfOrRpsCuutPwVs5z7oXjXbgi0PIZeaAM4dc87ZoiRy0LbrwG9MF61of7Mp4xj040sTjj0OO+7a+B2JRZ1CoAp/ZeU7PRgt4Ec4xHmK0g8r3PcMb23SE9bEILVBhbXD8iXv++acLrlt9rYnu7L3+8XbXeNXRxJqz/FVOfv+dpjMac6h5GzvCCw/csnjR9GKxZ9nGla++pxyVr5w9xhjOcT8w8Sztu093bNO29e22iNP0QwXW+6x+7qJ8rZYvwjQIC1NXflKNnsKCA34tbE3U1zdfYsifQuWb0SPL66uO+e3WXzx+8PjHvkq3du+yRlwcqAPlufatZbV8jzQVa/mV18mnIcSC09q1OSzgONgzmww8SwuVr6tHDteu+dmF++ePqNfh36M7Hzn4Dvyk+7xEQBXsfK5bXBuUYOhisbb71fBEdpxWTb1TnWLsf0wKjfvhI231EW05NrXV6/P/8DAA2T1BVagwfgchnyyrRWlwfTA/Z2JAZFMHUnNNsKOIN/KKuT7QKzTj6VKpzGHYrFyf//iqagKgQgWS8/SpSqiFaMXaxQLo7L1GKwCazktXxXl2GPVZqg25uyfXVRpkKtc/esdfoULlW4gzGCKoSwGQQx2p9wn//sFPzAUCAkVNYWv5UUUbsVK/fYIvUMWB09NTzB+LnZEDR7mdC+KNHDIHXFyhM4aLvBEgYl6a7Al0QgWGEi3qoFb5/E8JOdHc72Cv1j5Gyp5QH/LMWkvERULujDYTTi7nDfSmE2dw+odOdVo7NyhdVdAREdsdg6qQWR1UqbTzMXfIjcaL6lCH4vzJrk5r9NdVQ56FW6zDVaHL2dQBnrbS8U7bbu+FUUAfHEfutLYcEQLAFZLijYx7cJMthYKQc+GUTyVknV2h9q6Tu4cGJ/QDYPufkXagDvBX4oMz73fzbECA41ZnqkfuxNSR9YG2zrvBBlRZ+ks/nCNH9o880gzkheNjBRocRR87UGX8+26cNgi2I0HmZgTS13JDihPhIfP+yaQQwzH2O7T/acPcmUejlgLparmhxYnyGIES4ixlCsWBGE7PoFraqGt+Rh11NsgXmCJ2B5dXq5BXsIXq4O2tViEvdVJagZ5HC+SNg7kjJg3QkOPE9dEp5I3DcscINPQ4Op6YbXvhCGcTMDiPuBzFecB8g4BZ5FWz6YyaRpsmf3C72g9546jqBEDCFD4NHK0+asgxHG+jlvVRFPLDOZoQZwmqdzYmUMzf2ObQ5bzUUZ1NRgpCzi93suIY9AkVyogTKNTlW7NlUseQPzSHaD+UGUf0Q52eJWhGHKM+TKEG4DCgfWcPWUXNEkS8GvKHKnRDPy1BHVaARh13NjWH4Dm36r0uo4bc8bl92/6cGMYSDPLcwkPe3PRrqKgdOGYrCJAKHTDOVn2uVlTLNGVtcLqnFSy4OhyXk0H4so2HwJMBjcCBganqLDJxTt4MlAAn32MDsvK473d81MFxtp9fOYFUzUCDWIL6WMGCq2E0NRhoTKpPo9Qhdw/UYVzOCJQMxwpk1cdrrECqqIOMCRfooO7d84fTcTkDUFIcG5CNp3E4JRyXMylEc8fH2WiwMRcxhpyFp0FWAMEGo6kD5QEDUBocs0JmnsbhlMsDw3B8QQuUyNkkizcoZORpHM5wQAF59ECDU737ncjnRHogE0/DcOaXgAU+KqJEMYXS4xhCzsDTOCsAdQbK8wFJB5QFRw+k52mYOgMloEFtIN5oEikKZcPRAml5XDjuEhRG76EqoM5GIw2yB4ngWwJK52ySJ2gqBR2PC8e3yLl7AJwNvli00VfZFNI6mxVIw9M4HMwd6gQKElco71VRn9dNtmPNZpqiLhfnaRwOU4emDgLxNBL9kFcJCjhLVmqfTxCAEaAYz8wG3O/w3AEEydyCwGM5NBEeaHFW1IBDyOrFDqDnpDo7ytMonONz3NZoBqE2LH9QLaYQIevHWB/Ham2hOITscQDl7w2BIjxbG6QOObME0cbVoYEmAo5hYQ6Ry0x/lBg/h4cnGjmOE2hw2eoASOX5Ydn8QTxcVunznfCOILJEjXodqZ7ZxpShrwgBVDT8QoX2nddiebxMHtQ9165Q/gI9z4zJbpw3Lwk/QYig4Cr9GPzoxZAcoBAGGX4xS2CGQNdAOlQIPsEzAbWH6mBb7UDF6fDcBZtkfbqfHgFX0DhRdYidh6rzwNTaRX8JCgX+RpEULDgJBTrKABTBcQHlL9bx/NatDrzNNjZF1cFPEPIARBUSRhCEHWjFFCvTHNr1pO7Rqc72XfKFxsZaFSou+hIHkt52zf/qnzFigjF18F1mIKrOJ3ScDYEChRgUtwXQCY6IIUiBnhrdG3W5rt7R8q/A8JZaXS5/epxngjHS4AchDnyAvVb/JEml8E142nAZ682FQlghMKcTuqBCiMPuWMmBW/e2SFHX3tK791btX+OwAeXnxHkewrt8wyTjEHLbjbE/Gw9W0H8DfKB/7UZRnEgKcWUQAvAoFxWIKUR+c/L1T/b29sJ/KwGvL12/669526IzC1D+pirbO4y3VfvjD+gJurYyPJMjTetO6OtXHzgtdDRvwV8UeTk/XVQioUI8Zbi1MXlYEpUmH4OHXbVmw39fsXnz5rlHbFgjPWwonZEumnOouJE/th3wVGfsNPKo6uChm+56ra+/QzxhVmjur0x5H6/QH3dLY6BCodCxudEJD4fLVdr/Nm0qvFS7Cb/IYktsfu6fpcOL64bz4t+ey3YOeMhjw4dhomqm8rBTY0eGoDt04yXNHXTqe3HbDjZIPkdEGz2ZpJDi1jR9WBLl6t/THNu4afUyPVBx0Z4oz/F4KzkwLP5Vxvsd3dR9845fHjrrB/d97S7xW9hfWqSejio0sxzeBgksnMN3DsZ92h7THVu/beItQTTL6qA+MZ4zcRxGM5V3qrmjnkmNkA8VeeCUTCEofahD08yhSDyHMCBKH+1TD2lem3is6TOK4sYX2NvCeDtIB2JoFIiz4XxY2wrzCaI/uSDKQ4GwHxIo1OqYNtTqgMhbIDNOz+Cor0R5zhhOnZRVXOErVJF/H222aX3dxkE1DLhCCEQPjldLmIEgLJf+YDqeut2C05O/uIntHOrzc70+0I7SKeqBjWvabA1cjsHQUR7WCWHkYRxOvsN4ROkHgBO7WMGG4BYo5JkAl5BdRPmVXkVfoOu02SqA4LDwFSYSrtBNZR9HsKnTU9wt7oBCnjtuD+thdhr6ig7hq9D2weCCyQsBEMsYoEJro6906dGDkgyGRas6PfmVwpdCnu4VEHDBiQKN8CrmyiO8Qm57zA4YFgPC+6Egfegyu2DlNncCWdUBG/1EXIaQp7oVOyDmpuyVnR491VOhBWrvE2okFOLHpeehvgBLAzm3g7pwbmK9OUCFPOTwZBpweAoeDgGWp0LvydVOSANLCIT3Q7T9WG8zfWjlXS5/xM1JXOTo3IFTnL5AhJvMU/24xBKUOyo9JSDxNniE3AsXmQRi/RBUCtQO6BFFcIM+H42LEijr9tzp6am9HO4u6UO+uhME4v4Dp+Qo/Er6hFz1JkMCcYWqfExBhDUP6JK9x3ao01NbOVbPw07HT8IwKBTvyD1c7i0zT6BQAMMuGVzBkrUkdaqzmJc6FErWhyxcUY/oE/QW1BScIafvgHgmBTlEQxhzk0dAaWZ4fWNLTpxR58rRqvCQw5PEIFOoUpBNHqaw8FcWgQKFEEX6ypV/FqMINjiDbRS/8eHvUHnI+bfTuwZmAuJVzD1s+2gbDwOSx+VoyJXveTtofnQhKY7sb3isJgCCSwfngUiD1yAmGJOzY11iFShQSFKnnCt9HKUI1hPjRHkoEOjAWILUZd2sT8e6YLrRsjGNRA7h8TGQMfRyE4LeIwBhCx65E3lHjKfadP49paD5QiB6dliBO3KXy71leYpK2Db0Q4FCudLTkZgPmpgCJ8YDIXcMhpxkP4xOSOYCqtrK+ohCqFJ5/sMGedLgaHh4DgUmRLOIRx5GnCuHxv6D5T5FUoiWCnC7/XgjcXQ8hIBCLIUYlHBsuoY9kv1+yAeIzKTVdm6gvsLwH4wmtwIap7rYhRy6vST04XPqdrR3hSEZx+3DOk+F8MpMMtybpgo2INLxsByCswELIshodJMrh3wVypUmzdBHW1ocAw+posuJ2odaAYUCOC+X8wGqzizXJ80IzExZSI1j4hE5RCXi3kphuFZuhayFAuuHDu5vtDqmeIOLhf0QNp4BMEtgGkEq+bicC+g6fOhXO6VXx8LDFEJpgILfTtLAo2p5uNxJHrbdcBxzvLFajqcOdTfFttHlGmDbGqAs6lj1IdVj7sF+KFSF09HizhFy17xNfEwhzrMkPmYMnXA41SI3CJEj6P2a70RrOVaaYtCxJYw32hHZFOr++MoZTW6gyB9xgdMuucAepg4cW7whFFeImYJIJmES5uK0+8IR9StnkLEnuU2BXzo+y4rj4kGFuCy8YxVoKJhJIcAp50ZA55IUKDOOi4e6nEgbJGDZRKGwxNOWPhQnx4CcpY8cctlxnDzsfojlDKsO0CCoSVA2jctxHF8gGGhsVLDBcax+QM9DXU6kDPVsNAZQhlNFi9MAJylQA9Tx4aGVAjQeEwYR5ASiH+ao/ZCEw4GcHSsLuYbgeOgDIuEdK9ODweBrsKSagoIDQN4u58YRj+jw6NTP3PEG7+NjClwZ7gmgFUsoGSiC4w80NmO/I+i8eJhC2H4pcWgKcSIRcjEcEXKufmi7Y9TB1Y0KHL94YwrR/KeRplRy6Nrc5TQ4Asje6w9OtwN74/jyUIXke1U5gZCIKoTdaHyiHaujH7IO2vX44/jz8JFTmjSQSBBo7BVYMJMQ6CEtjsghu0JhwRlfSoDjz0Po2DbLFwg68UU9IpejA43G5818XC6OIbYkwUnCwxSSteEwVJ94nElbfGo50fzoPBFOEh7IIRwkwS8aYyLgxCap/bHF0l+94yxOoyBsvTbKq98RBufp12z3JgDiWUM/2eVRxzdpHs0SD2vlSvfAX4sc66gUdEDJ1PGqdwQ6znHkFOtq4QbCuUU6GeZ8nC05UEJ1EvPwSkFtNkKhRtjZav8BDrsoSYES4yTLH2wU+wSPIbBAg1f7lzQK6rrBU0MuOU5yHjamEAWgASgg5TnoFaiDlyMJUAqc5DygEIz6KPlDOyWjSqVJh6vS8x/+QGlw0vCgy0X04TAIJn1jOKI6Ek0ChVLhpOGhOUQrbeFz2B0xJHwV3zSrpNzBcMPJzxTS4aTjUXKIgYgAlKwBzU7JHYbjB5QSJx1P6HIivEIoWKJq0S2xYGNIboXS4qTlURSimighx7bo1WEh57jf+XOiIkcID/NE9Y70vojL8TBDM+D6wFyTO+IIdpdLrU4GHq4QD7QwxLgxUGc7LJofn9tCLgNOen1EDgln48pQhUTuyP1OFAmADDeltdTBBudIHW981IcaAq3eBBgagsnZZKglt+QH1fKGrhVrK+Xn2eR3+Cxn4GEKYbrwsVLe91Cl6pMOq92opjGv7o5JVKxNf3msZlfvTZl4ms6nz8uhMqhKqFWuvt/wybXSsPdOKip/Cn+wlr9pu7JH4pVMPIQcfgT/7DpVhJZAzB7K9RV36J8riLZv+7FX1+A/XsgP5uG/XajtXnl68OBxdE/P9Yw81VWPz6/jU7RcI/ydAfi7+L/Y54cDjVz94ZxfbVy0bNHfjLr43j96NtqyW0Ye+C2xCU+X62JYAT+xq89/92GbscUaUyXX7lm95wVnusXeqNuQmQd8bsKFk8p1PpUnnQK/af7ZTQ3gAaJZE7Y+9O6KFe+ecuqEWd6RNiTQDeEh1SqBEGuqNsF8SJrpfdDG8Hifbsh3/IJnyC9xphN83vQh5Pnnn1++fPlffD6m5cv/D5ZcYva0NJiSAAAAAElFTkSuQmCC"
            alt="IFC Icon"
            style={{ width: 16, height: 16 }}
          />
        );
      },
      onClick: onLoadIFC,
    },
  ];

  const columns: IColumn[] = [
    {
      key: "name",
      name: "Name",
      fieldName: "Name",
      minWidth: 140,
      maxWidth: 250,
      isResizable: true,
      onRender: (item: IDocument) => {
        return (
          <a href="#" onClick={() => openDocument(item.FileRef)}>
            {item.Name}
          </a>
        );
      },
    },
    {
      key: "modified",
      name: "Modified",
      fieldName: "Modified",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    {
      key: "modifiedBy",
      name: "Modified By",
      fieldName: "ModifiedBy",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
    },
    {
      key: "actions",
      name: "Actions",
      minWidth: 70,
      maxWidth: 70,
      onRender: (item: IDocument) => {
        const menuProps: IContextualMenuProps = {
          items: [
            {
              key: "open",
              text: "Open",
              iconProps: { iconName: "OpenFile" },
              onClick: () => openDocument(item.FileRef),
            },
            {
              key: "loadIFC",
              text: "Load to IFC Viewer",
              onRenderIcon: () => {
                return (
                  <img
                    src="data:image/png;base64,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"
                    alt="IFC Icon"
                    style={{ width: 16, height: 16 }}
                  />
                );
              },
              onClick: () => onLoadIFC(),
            },
          ],
        };

        return (
          <IconButton
            menuProps={menuProps}
            iconProps={{ iconName: "MoreVertical" }}
          />
        );
      },
    },
  ];

  const { isLoading, error, documents } = state;

  return (
    <div className={styles.documentLibraryViewer}>
      <div className={styles.header}>
        <Text variant="large">{props.title || "Document Library"}</Text>
      </div>

      <CommandBar items={commandItems} />

      {isLoading && <div>Loading documents...</div>}
      {error && <div className={styles.error}>Error: {error}</div>}

      {!isLoading && !error && (
        <DetailsList
          items={documents}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.multiple}
          isHeaderVisible={true}
          className={styles.documentsList}
        />
      )}
    </div>
  );
};

export default DocumentLibraryViewer;
